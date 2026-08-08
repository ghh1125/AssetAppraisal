from pathlib import Path

from openpyxl import Workbook

from demo.adapters.semantic_excel import _income_label, extract_workbook_facts


def _save_workbook(path: Path, sheets: dict[str, list[list[object]]]) -> Path:
    workbook = Workbook()
    workbook.remove(workbook.active)
    for title, rows in sheets.items():
        sheet = workbook.create_sheet(title)
        for row in rows:
            sheet.append(row)
    workbook.save(path)
    return path


def test_asset_summary_is_matched_by_labels_and_normalized_to_wan(tmp_path: Path):
    path = _save_workbook(
        tmp_path / "arbitrary-name.xlsx",
        {
            "资产评估结果分类汇总表（元）": [
                ["金额单位：人民币元"],
                ["序号", "科目名称", "账面价值", "评估价值", "增减值", "增值率%"],
                [1, "七、所有者权益（净资产）", 86_979_689.29, 93_972_005.69, 6_992_316.4, 8.04],
            ]
        },
    )

    facts = extract_workbook_facts(path, "reporting_workbook")

    assert facts["fields"]["book_net_assets"] == 8697.968929
    assert facts["fields"]["asset_approach_value"] == 9397.200569
    assert facts["evidence"]["asset_approach_value"]["locator"].endswith("!D3")


def test_asset_scope_uses_summary_title_inside_a_generic_sheet_name(tmp_path: Path):
    """Asset-cleanup workbooks commonly name sheets 表1 rather than 汇总表."""
    path = _save_workbook(
        tmp_path / "uploaded-asset-inventory.xlsx",
        {
            "表1": [
                ["资产评估结果--汇总表"],
                ["金额单位：人民币万元"],
                ["序号", "项目", "账面价值", "评估价值"],
                [1, "流动资产", 14853.73, 14853.73],
                [2, "非流动资产", 1518.19, 3285.07],
                [3, "其中：固定资产净额", 499.36, 856.24],
                [4, "无形资产净额", 0, 500],
                [5, "长期待摊费用", 178.83, 178.83],
                [6, "资产总计", 16371.92, 18138.8],
                [7, "流动负债", 11773.76, 11773.76],
                [8, "非流动负债", 0, 0],
                [9, "负债总计", 11773.76, 11773.76],
                [10, "净资产（所有者权益）", 4598.16, 6365.04],
            ]
        },
    )

    facts = extract_workbook_facts(path, "reporting_workbook")

    rows = facts["fields"]["asset_scope_summary_table"]["rows"]
    assert rows[0] == ["流动资产账面金额：", "148,537,300.00"]
    assert rows[2] == ["其中：固定资产账面金额：", "4,993,600.00"]
    assert rows[8] == ["资产合计账面金额：", "163,719,200.00"]
    assert facts["evidence"]["asset_scope_summary_table"]["locator"].startswith("表1!")


def test_all_zero_appraisal_column_keeps_asset_result_unresolved(tmp_path: Path):
    path = _save_workbook(
        tmp_path / "unfinished-appraisal.xlsx",
        {
            "1-汇总表": [
                ["金额单位：人民币万元"],
                ["项目", "账面价值", "评估价值"],
                ["流动资产", 100, 0],
                ["非流动资产", 200, 0],
                ["负债合计", 70, 0],
                ["净资产", 230, 0],
            ]
        },
    )

    facts = extract_workbook_facts(path, "reporting_workbook")

    assert facts["fields"]["book_net_assets"] == 230
    assert "asset_approach_value" not in facts["fields"]
    assert facts["evidence"]["asset_approach_value"] == {
        "kind": "unfinished_appraisal",
        "file": path.name,
        "locator": "1-汇总表!C6",
    }
    assert any("[unfinished_appraisal]" in issue for issue in facts["issues"])


def test_net_profit_label_accepts_accounting_loss_suffix():
    assert (
        _income_label("五、净利润（净亏损以“-”号填列）")
        == "四、净利润"
    )


def test_zero_net_result_remains_valid_when_appraisal_column_has_activity(
    tmp_path: Path,
):
    path = _save_workbook(
        tmp_path / "valid-zero-appraisal.xlsx",
        {
            "汇总表": [
                ["金额单位：人民币万元"],
                ["项目", "账面价值", "评估价值"],
                ["资产总计", 100, 80],
                ["负债合计", 50, 80],
                ["净资产", 50, 0],
            ]
        },
    )

    facts = extract_workbook_facts(path, "reporting_workbook")

    assert facts["fields"]["asset_approach_value"] == 0
    assert not any("[unfinished_appraisal]" in issue for issue in facts["issues"])


def test_income_value_uses_equity_label_and_sheet_unit(tmp_path: Path):
    path = _save_workbook(
        tmp_path / "income-any-name.xlsx",
        {
            "净现金流计算表": [
                ["金额单位：元"],
                ["企业整体价值", None, 69_000_164.71],
                ["股东全部权益价值", None, 68_500_000],
            ],
            "主要产品及服务": [
                ["主要产品及服务"],
                ["产品或服务名称", "增值税率（%）"],
                ["滤波器", 0.13],
                ["天线", 0.13],
            ],
        },
    )

    facts = extract_workbook_facts(path, "income_workbook")

    assert facts["fields"]["income_approach_value"] == 6850
    assert facts["fields"]["main_products"] == "主要产品及服务：滤波器、天线。"
    assert facts["fields"]["tax_rates"].startswith("被评估单位执行《企业会计准则》")
    assert "13%" in facts["fields"]["tax_rates"]


def test_market_value_falls_back_to_net_asset_evaluation_row(tmp_path: Path):
    path = _save_workbook(
        tmp_path / "market.xlsx",
        {
            "1-汇总表": [
                ["评估方法", "市场法"],
                ["金额单位：人民币万元"],
                ["项目", "序号", "账面价值", "评估价值"],
                ["净资产", 15, 29_151.74, 101_000],
            ]
        },
    )

    facts = extract_workbook_facts(path, "income_workbook")

    assert facts["fields"]["market_approach_value"] == 101_000
    assert "income_approach_value" not in facts["fields"]


def test_market_result_label_can_be_generic_conclusion(tmp_path: Path):
    path = _save_workbook(
        tmp_path / "market-conclusion.xlsx",
        {
            "估值过程和结果": [
                ["采用市场法"],
                ["项目", "金额"],
                ["评估结论", 9000],
            ]
        },
    )
    facts = extract_workbook_facts(path, "income_workbook")
    assert facts["fields"]["market_approach_value"] == 9000
    assert facts["evidence"]["market_approach_value"]["locator"].endswith("!B3")


def test_summary_header_can_be_far_above_net_asset_total(tmp_path: Path):
    rows = [
        ["金额单位：人民币万元"],
        ["项目", "序号", "账面价值", "评估价值"],
    ]
    rows.extend([[f"资产科目{index}", index, index * 10, index * 11] for index in range(1, 24)])
    rows.append(["净 资 产（所有者权益）", 24, 21_628.2, 22_249.73])
    path = _save_workbook(tmp_path / "long-summary.xlsx", {"汇总表": rows})

    facts = extract_workbook_facts(path, "reporting_workbook")

    assert facts["fields"]["book_net_assets"] == 21_628.2
    assert facts["fields"]["asset_approach_value"] == 22_249.73


def test_legacy_accounting_header_uses_zhang_variant(tmp_path: Path):
    path = _save_workbook(
        tmp_path / "legacy-income.xlsx",
        {
            "结果汇总": [
                ["评估方法", "收益法"],
                ["金额单位：人民币万元"],
                ["项目", "序号", "帐面价值", "调整后帐面值", "评估值", "评估增值"],
                ["净 资 产", 14, 1_382.65, 1_382.65, 2_800, 1_417.35],
            ]
        },
    )

    facts = extract_workbook_facts(path, "income_workbook")

    assert facts["fields"]["income_approach_value"] == 2_800
    assert "market_approach_value" not in facts["fields"]


def test_asset_summary_builds_scope_and_long_term_tables(tmp_path: Path):
    path = _save_workbook(
        tmp_path / "asset-details.xlsx",
        {
            "汇总表": [
                ["金额单位：人民币万元"],
                ["项目", "序号", "账面价值", "评估价值"],
                ["流动资产", 1, 100, 101],
                ["非流动资产", 2, 200, 210],
                ["其中：固定资产净额", 3, 120, 130],
                ["其中：无形资产净额", 4, 30, 35],
                ["长期待摊费用", 5, 10, 10],
                ["资产总计", 6, 300, 311],
                ["流动负债", 7, 50, 50],
                ["非流动负债", 8, 20, 20],
                ["负债总计", 9, 70, 70],
                ["净资产", 10, 230, 241],
            ],
            "固定资产汇总表": [
                ["金额单位：人民币元"],
                ["项目", "账面价值", "评估价值"],
                ["电子设备", 320_000, 350_000],
            ],
        },
    )

    fields = extract_workbook_facts(path, "reporting_workbook")["fields"]

    scope = fields["asset_scope_summary_table"]["rows"]
    assert ["流动资产账面金额：", "1,000,000.00"] in scope
    assert ["所有者权益账面金额：", "2,300,000.00"] in scope
    long_term = fields["long_term_assets_table"]["rows"]
    assert long_term[1][:2] == ["电子设备", "320,000.00"]
    assert "固定资产账面价值1,200,000.00元" in fields["major_long_term_assets"]


def test_historical_financial_tables_follow_labels_not_coordinates(tmp_path: Path):
    path = _save_workbook(
        tmp_path / "financial-history.xlsx",
        {
            "资产负债表": [
                ["金额单位：人民币元"],
                ["项目", "2023年", "2024年", "2025年6月30日"],
                ["资产总计", 100, 200, 300],
                ["负债合计", 40, 70, 90],
                ["所有者权益合计", 60, 130, 210],
            ],
            "利润表": [
                ["金额单位：人民币元"],
                ["项目", "2023年度", "2024年度", "2025年1-6月"],
                ["一、营业收入", 80, 120, 150],
                ["减：营业成本", 30, 50, 60],
                ["四、净利润", 20, 30, 40],
            ],
        },
    )

    fields = extract_workbook_facts(path, "audited_financials")["fields"]

    balance = fields["historical_balance_sheet_table"]["rows"]
    assert balance[0] == ["项目\\报表日", "2023年", "2024年", "2025年6月30日"]
    assert balance[1] == ["总资产", "100.00", "200.00", "300.00"]
    income = fields["historical_income_statement_table"]["rows"]
    assert income[1] == ["一、营业收入", "80.00", "120.00", "150.00"]
    assert income[-1] == ["四、净利润", "20.00", "30.00", "40.00"]


def test_dual_sided_balance_ignores_the_other_section_sequence_column(tmp_path: Path):
    path = _save_workbook(
        tmp_path / "dual-balance.xlsx",
        {
            "资产负债表": [
                ["金额单位：人民币元"],
                ["资产", "期初数", "期末数", "序号", "负债及所有者权益", "期初数", "期末数"],
                ["资产总计", 100, 200, 88, "负债及所有者权益合计", 100, 200],
                [None, None, None, None, "负债合计", 40, 70],
                [None, None, None, None, "所有者权益合计", 60, 130],
            ]
        },
    )

    rows = extract_workbook_facts(path, "audited_financials")["fields"][
        "historical_balance_sheet_table"
    ]["rows"]

    assert rows[1] == ["总资产", "XXX", "100.00", "200.00"]


def test_income_workbook_can_supply_history_when_asset_file_has_none(tmp_path: Path):
    path = _save_workbook(
        tmp_path / "income-history.xlsx",
        {
            "历资表": [
                ["金额单位：人民币万元"],
                ["项目", "2023年", "2024年", "2025年"],
                ["资产总计", 1, 2, 3],
                ["负债合计", 0.4, 0.7, 0.9],
                ["所有者权益合计", 0.6, 1.3, 2.1],
            ],
            "历利表": [
                ["金额单位：人民币万元"],
                ["项目", "2023年", "2024年", "2025年"],
                ["营业收入", 0.8, 1.2, 1.5],
                ["净利润", 0.2, 0.3, 0.4],
            ],
        },
    )

    fields = extract_workbook_facts(path, "income_workbook")["fields"]

    assert fields["historical_balance_sheet_table"]["rows"][1][-1] == "30,000.00"
    assert fields["historical_income_statement_table"]["rows"][-1][-1] == "4,000.00"


def test_history_uses_period_columns_and_derives_missing_equity(tmp_path: Path):
    path = _save_workbook(
        tmp_path / "multi-header.xlsx",
        {
            "资产负债表": [
                ["金额单位：人民币元"],
                ["项目", "2023年度", "2024年度", "2024年度", "2025年度"],
                [None, "实际数", "审定数", "增长率%", "预测数"],
                ["资产总计", 100, 200, 1.0, 300],
                ["负债合计", 40, 70, 0.75, 100],
            ]
        },
    )

    facts = extract_workbook_facts(path, "audited_financials")
    rows = facts["fields"]["historical_balance_sheet_table"]["rows"]

    assert rows[0] == ["项目\\报表日", "历史期1", "2023年度", "2024年度"]
    assert rows[1] == ["总资产", "XXX", "100.00", "200.00"]
    assert rows[3] == ["所有者权益", "XXX", "60.00", "130.00"]
    assert facts["evidence"]["historical_balance_sheet_table"]["kind"] == (
        "semantic_excel_derived"
    )


def test_numeric_amount_never_becomes_a_historical_period_header(tmp_path: Path):
    path = _save_workbook(
        tmp_path / "numeric-header-trap.xlsx",
        {
            "利润表": [
                ["金额单位：人民币万元"],
                ["项目", "2022年度", "2023年度", "2024年度"],
                ["营业收入", 59_734.507071, 56_129.203566, 58_991.419895],
                ["净利润", 100, 200, 300],
            ]
        },
    )

    rows = extract_workbook_facts(path, "audited_financials")["fields"][
        "historical_income_statement_table"
    ]["rows"]

    assert rows[0] == ["项目\\报表年度", "2022年度", "2023年度", "2024年度"]


def test_history_keeps_valuation_period_after_calendar_years(tmp_path: Path):
    path = _save_workbook(
        tmp_path / "valuation-period.xlsx",
        {
            "利润表": [
                ["金额单位：人民币元"],
                ["项目", "2023年", "2024年", "评估基准期"],
                ["营业收入", 100, 200, 300],
                ["净利润", 10, 20, 30],
            ]
        },
    )

    rows = extract_workbook_facts(path, "audited_financials")["fields"][
        "historical_income_statement_table"
    ]["rows"]

    assert rows[0] == ["项目\\报表年度", "2023年", "2024年", "评估基准期"]
    assert rows[1] == ["一、营业收入", "100.00", "200.00", "300.00"]


def test_long_term_assets_fall_back_to_detail_rows(tmp_path: Path):
    path = _save_workbook(
        tmp_path / "asset-detail-only.xlsx",
        {
            "固定资产明细表": [
                ["金额单位：人民币元"],
                [
                    "资产编号",
                    "资产名称",
                    "资产类别",
                    "账面原值",
                    "累计折旧",
                    "账面净值",
                    "评估值",
                ],
                ["D001", "电脑", "办公电子设备", 10_000, 2_000, 8_000, 8_500],
                ["D002", "打印机", "电子设备类", 5_000, 1_000, 4_000, 4_200],
                ["", "合计", "", 15_000, 3_000, 12_000, 12_700],
            ],
            "汇总表": [
                ["金额单位：人民币元"],
                ["项目", "账面价值", "评估价值"],
                ["无形资产", 30_000, 35_000],
                ["长期待摊费用", 10_000, 10_000],
            ],
        },
    )

    facts = extract_workbook_facts(path, "reporting_workbook")
    rows = facts["fields"]["long_term_assets_table"]["rows"]

    assert rows[1] == ["电子设备", "12,000.00", "2项", "以评估明细表为准"]
    assert "固定资产明细表!F3:F4" in facts["evidence"][
        "long_term_assets_table"
    ]["locator"]


def test_electronic_summary_accepts_prefixed_category_label(tmp_path: Path):
    path = _save_workbook(
        tmp_path / "prefixed-electronics.xlsx",
        {
            "其他资产汇总表": [
                ["金额单位：人民币元"],
                ["项目", "账面原值", "账面净值", "评估原值", "评估净值"],
                ["固定资产—电子设备", 15_000, 12_000, 16_000, 13_000],
                ["无形资产", 30_000, 30_000, 35_000, 35_000],
            ]
        },
    )

    rows = extract_workbook_facts(path, "reporting_workbook")["fields"][
        "long_term_assets_table"
    ]["rows"]

    assert rows[1][1] == "12,000.00"


def test_electronic_detail_sheet_supports_two_row_headers(tmp_path: Path):
    path = _save_workbook(
        tmp_path / "two-row-electronics.xlsx",
        {
            "4-8-7电子设备": [
                ["固定资产—电子设备评估明细表", "金额单位：人民币元"],
                [
                    "序号",
                    "资产编号",
                    "设备名称",
                    "数量",
                    "账面价值",
                    None,
                    "评估价值",
                    None,
                ],
                [None, None, None, None, "原值", "净值", "原值", "净值"],
                [1, "D001", "电脑", 1, 10_000, 8_000, 10_500, 8_500],
                [2, "D002", "打印机", 1, 5_000, 4_000, 5_200, 4_200],
                [None, None, "合计", 2, 15_000, 12_000, 15_700, 12_700],
            ],
            "汇总表": [
                ["金额单位：人民币元"],
                ["项目", "账面价值", "评估价值"],
                ["无形资产", 30_000, 35_000],
            ],
        },
    )

    facts = extract_workbook_facts(path, "reporting_workbook")
    rows = facts["fields"]["long_term_assets_table"]["rows"]

    assert rows[1] == ["电子设备", "12,000.00", "2项", "以评估明细表为准"]
    assert "4-8-7电子设备!F4:F5" in facts["evidence"][
        "long_term_assets_table"
    ]["locator"]


def test_electronic_summary_uses_book_net_subcolumn(tmp_path: Path):
    path = _save_workbook(
        tmp_path / "two-row-summary.xlsx",
        {
            "固定资产汇总表": [
                ["金额单位：人民币元"],
                ["项目", "账面价值", None, "评估价值", None],
                [None, "原值", "净值", "原值", "净值"],
                ["固定资产—电子设备", 15_000, 12_000, 16_000, 13_000],
            ],
            "汇总表": [
                ["金额单位：人民币元"],
                ["项目", "账面价值", "评估价值"],
                ["无形资产", 30_000, 35_000],
            ],
        },
    )

    facts = extract_workbook_facts(path, "reporting_workbook")
    rows = facts["fields"]["long_term_assets_table"]["rows"]

    assert rows[1][1] == "12,000.00"
    assert facts["evidence"]["long_term_assets_table"]["locator"].endswith(
        "固定资产汇总表!C4"
    )


def test_ambiguous_electronic_book_columns_remain_unresolved(tmp_path: Path):
    path = _save_workbook(
        tmp_path / "ambiguous-electronics.xlsx",
        {
            "电子设备": [
                ["固定资产—电子设备明细表", "金额单位：人民币元"],
                ["资产编号", "资产名称", "账面净值", "账面净额"],
                ["D001", "电脑", 8_000, 7_500],
            ],
            "汇总表": [
                ["金额单位：人民币元"],
                ["项目", "账面价值", "评估价值"],
                ["无形资产", 30_000, 35_000],
            ],
        },
    )

    facts = extract_workbook_facts(path, "reporting_workbook")

    assert facts["fields"]["long_term_assets_table"]["rows"][1][1] == "XXX"
    assert any(
        "long_term_assets_table" in issue
        and "[ambiguous_candidate]" in issue
        for issue in facts["issues"]
    )
