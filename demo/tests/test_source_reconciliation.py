from demo.domain.source_reconciliation import (
    SourceCandidate,
    reconcile_field,
)
from demo.pipeline import _reviewable_source_conflicts, _reviewable_source_fallbacks


def test_pdf_authoritative_field_keeps_pdf_value_and_reports_excel_difference():
    result = reconcile_field(
        field_key="historical_balance_sheet_table",
        candidates=[
            SourceCandidate(
                value=163_719_131.79,
                source_kind="pdf_ocr",
                source_file="审计报告.pdf",
                source_locator="第58页 资产负债表",
            ),
            SourceCandidate(
                value=169_478_239.93,
                source_kind="income_workbook",
                source_file="收益法.xlsx",
                source_locator="历资表!D12",
            ),
        ],
        source_priority=("pdf_ocr", "asset_workbook", "income_workbook"),
        require_primary_source=True,
    )

    assert result.value == 163_719_131.79
    assert result.source_kind == "pdf_ocr"
    assert len(result.discrepancies) == 1
    assert result.discrepancies[0].other_source_file == "收益法.xlsx"


def test_pdf_authoritative_field_stays_unresolved_when_pdf_is_absent():
    result = reconcile_field(
        field_key="historical_income_statement_table",
        candidates=[
            SourceCandidate(
                value=21_652_412.25,
                source_kind="income_workbook",
                source_file="收益法.xlsx",
                source_locator="历利表!D8",
            )
        ],
        source_priority=("pdf_ocr", "asset_workbook", "income_workbook"),
        require_primary_source=True,
    )

    assert result.value is None
    assert result.source_kind == "missing"
    assert result.issues == ["historical_income_statement_table：未上传审计PDF，无法获取审计数据"]


def test_pdf_field_uses_workbook_when_pdf_is_absent_if_fallback_is_enabled():
    result = reconcile_field(
        field_key="historical_income_statement_table",
        candidates=[
            SourceCandidate(
                21_652_412.25,
                "income_workbook",
                "收益法.xlsx",
                "历利表!D8",
            )
        ],
        source_priority=("pdf_ocr", "income_workbook"),
        require_primary_source=True,
        allow_fallback_when_primary_missing=True,
    )

    assert result.value == 21_652_412.25
    assert result.source_kind == "income_workbook"
    assert "暂未完成PDF对照" in result.issues[0]


def test_pdf_uploaded_but_ocr_missed_field_says_so_when_using_excel_fallback():
    result = reconcile_field(
        field_key="historical_income_statement_table",
        candidates=[
            SourceCandidate(
                21_652_412.25,
                "income_workbook",
                "收益法.xlsx",
                "历利表!D8",
            )
        ],
        source_priority=("pdf_ocr", "income_workbook"),
        require_primary_source=True,
        allow_fallback_when_primary_missing=True,
        primary_source_supplied=True,
    )

    assert result.value == 21_652_412.25
    assert result.issues == [
        "historical_income_statement_table：审计PDF已上传，但OCR未识别到该字段；"
        "已采用收益法.xlsx（历利表!D8），暂未完成PDF对照"
    ]


def test_excel_fallback_produces_a_non_conflict_word_note():
    findings = _reviewable_source_fallbacks(
        [
            {
                "field_key": "book_net_assets",
                "selected": {
                    "value": "100.00",
                    "source_kind": "asset_workbook",
                    "source_file": "资产基础法.xlsx",
                    "source_locator": "汇总表!C8",
                },
            }
        ],
        {"book_net_assets": "账面净资产"},
    )

    assert findings == [
        {
            "review_kind": "excel_fallback",
            "field_key": "book_net_assets",
            "field_name": "账面净资产",
            "excel_value": "100.00",
                "excel_file": "资产基础法.xlsx",
                    "excel_locator": "《资产基础法.xlsx》工作表“汇总表”（账面净资产）",
                "pdf_uploaded": False,
                "word_context_hint": "",
        }
    ]


def test_selected_valuation_method_can_use_its_configured_workbook():
    result = reconcile_field(
        field_key="final_appraisal_value",
        candidates=[
            SourceCandidate(
                value=6365.04,
                source_kind="asset_workbook",
                source_file="资产基础法.xlsx",
                source_locator="结果汇总!D31",
            )
        ],
        source_priority=("asset_workbook",),
        require_primary_source=True,
    )

    assert result.value == 6365.04
    assert result.discrepancies == []


def test_equal_numeric_values_with_different_number_formats_are_deduplicated():
    result = reconcile_field(
        field_key="book_net_assets",
        candidates=[
            SourceCandidate("4,598.16", "pdf_ocr", "审计报告.pdf", "第58页"),
            SourceCandidate(4598.1600, "asset_workbook", "资产基础法.xlsx", "表1!C31"),
        ],
        source_priority=("pdf_ocr", "asset_workbook"),
        require_primary_source=True,
    )

    assert result.value == "4,598.16"
    assert result.discrepancies == []


def test_equal_table_values_with_different_number_formats_are_deduplicated():
    result = reconcile_field(
        field_key="historical_income_statement_table",
        candidates=[
            SourceCandidate(
                {"rows": [["项目", "2024年度"], ["营业收入", "1,000.00"]]},
                "pdf_ocr",
                "审计报告.pdf",
                "第58页",
            ),
            SourceCandidate(
                {"rows": [["项目", "2024年度"], ["营业收入", 1000]]},
                "income_workbook",
                "收益法.xlsx",
                "历利表!D8",
            ),
        ],
        source_priority=("pdf_ocr", "income_workbook"),
        require_primary_source=True,
    )

    assert result.discrepancies == []


def test_partial_pdf_table_keeps_pdf_rows_and_fills_only_missing_rows_from_excel():
    result = reconcile_field(
        field_key="asset_scope_summary_table",
        candidates=[
            SourceCandidate(
                {
                    "rows": [
                        ["项目", "账面金额（元）"],
                        ["流动资产账面金额：", "100.00"],
                        ["固定资产账面金额：", "XXX"],
                    ]
                },
                "pdf_ocr",
                "审计报告.pdf",
                "OCR结构化结果.xlsx / asset_scope_summary_table",
            ),
            SourceCandidate(
                {
                    "rows": [
                        ["项目", "账面金额（元）"],
                        ["流动资产账面金额：", "90.00"],
                        ["固定资产账面金额：", "50.00"],
                    ]
                },
                "asset_workbook",
                "资产基础法.xlsx",
                "汇总表!B3:B4",
            ),
        ],
        source_priority=("pdf_ocr", "asset_workbook"),
        require_primary_source=True,
    )

    assert result.source_kind == "pdf_ocr"
    assert result.value["rows"] == [
        ["项目", "账面金额（元）"],
        ["流动资产账面金额：", "100.00"],
        ["固定资产账面金额：", "50.00"],
    ]
    assert len(result.discrepancies) == 1


def test_partial_pdf_summary_without_a_header_fills_its_first_missing_row_from_excel():
    result = reconcile_field(
        field_key="asset_scope_summary_table",
        candidates=[
            SourceCandidate(
                {"rows": [["流动资产账面金额：", "XXX"], ["资产合计账面金额：", "100.00"]]},
                "pdf_ocr",
                "审计报告.pdf",
                "OCR结构化结果.xlsx / asset_scope_summary_table",
            ),
            SourceCandidate(
                {"rows": [["流动资产账面金额：", "90.00"], ["资产合计账面金额：", "100.00"]]},
                "asset_workbook",
                "资产基础法.xlsx",
                "汇总表!B3:B4",
            ),
        ],
        source_priority=("pdf_ocr", "asset_workbook"),
        require_primary_source=True,
    )

    assert result.value["rows"][0] == ["流动资产账面金额：", "90.00"]


def test_reviewable_table_conflict_keeps_row_and_period_for_word_annotation():
    rows = [
        {
            "field_key": "historical_income_statement_table",
            "discrepancies": [
                {
                    "selected_source_file": "审计报告.pdf",
                    "selected_source_locator": "第58页 利润表",
                    "selected_value": {
                        "rows": [
                            ["项目", "2024年度"],
                            ["一、营业收入", "100.00"],
                        ]
                    },
                    "other_source_file": "收益法.xlsx",
                    "other_source_locator": "历利表!D8",
                    "other_value": {
                        "rows": [
                            ["项目", "2024年度"],
                            ["一、营业收入", "90.00"],
                        ]
                    },
                }
            ],
        }
    ]

    findings = _reviewable_source_conflicts(
        rows,
        {"historical_income_statement_table": "历史利润表"},
    )

    assert findings == [
        {
            "field_key": "historical_income_statement_table",
            "field_name": "历史利润表 / 一、营业收入 / 2024年度",
            "pdf_value": "100.00",
            "excel_value": "90.00",
            "word_context_hint": "一、营业收入",
            "pdf_file": "审计报告.pdf",
                "pdf_locator": "审计 PDF 第58页 利润表：历史利润表",
                "excel_file": "收益法.xlsx",
                "excel_locator": "《收益法.xlsx》工作表“历利表”（历史利润表）",
            }
        ]
