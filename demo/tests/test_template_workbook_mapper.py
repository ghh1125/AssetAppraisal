from demo.adapters.audit_intake import ExtractedTable
from openpyxl import Workbook

from demo.adapters.template_workbook_mapper import (
    SourceRow,
    _asset_target_accepts_statement_value,
    _find_source,
)


def test_find_source_falls_back_to_unknown_scope_only_when_preferred_scope_missing():
    table = ExtractedTable(
        source_file="审计报告.pdf",
        page_number=1,
        table_id="t1",
        category="利润表",
        roles=("利润表",),
        matrix=[],
        period_headers={},
        confidence="high",
    )
    source_rows = {
        "营业收入": [
            SourceRow(table, 2, "营业收入", ((2, 1000.0),), ((2, "2025年度"),), 1.0, "未识别"),
        ]
    }

    source = _find_source("营业收入", source_rows, preferred_scope="母公司")

    assert source is not None
    assert source.scope == "未识别"


def test_find_source_uses_requested_scope_when_both_explicit_scopes_exist():
    table = ExtractedTable(
        source_file="审计报告.pdf",
        page_number=1,
        table_id="t1",
        category="利润表",
        roles=("利润表",),
        matrix=[],
        period_headers={},
        confidence="high",
    )
    source_rows = {
        "营业收入": [
            SourceRow(table, 2, "营业收入", ((2, 1000.0),), ((2, "2025年度"),), 1.0, "母公司"),
            SourceRow(table, 3, "营业收入", ((2, 2000.0),), ((2, "2025年度"),), 1.0, "合并"),
        ]
    }

    source = _find_source("营业收入", source_rows, preferred_scope="母公司")

    assert source is not None
    assert source.scope == "母公司"


def test_asset_statement_value_is_not_written_into_original_cost_column():
    workbook = Workbook()
    sheet = workbook.active
    sheet["D7"] = "账面原值"
    sheet["C8"] = "固定资产合计"

    assert not _asset_target_accepts_statement_value(sheet, 8, 4)


def test_asset_statement_value_maps_only_to_net_row_of_gross_allowance_net_block():
    workbook = Workbook()
    sheet = workbook.active
    sheet["D7"] = "账面价值"
    sheet["C8"] = "存货合计"
    sheet["C9"] = "减：存货跌价准备"
    sheet["C10"] = "存货净额"

    assert not _asset_target_accepts_statement_value(sheet, 8, 4)
    assert not _asset_target_accepts_statement_value(sheet, 9, 4)
    assert _asset_target_accepts_statement_value(sheet, 10, 4)
