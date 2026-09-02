from openpyxl import Workbook

from demo.adapters.note_evidence_mapper import _append_mapping


def test_more_specific_note_mapping_retires_prior_effective_mapping():
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "表3-9"
    cell = sheet["D16"]
    mappings = [[
        "表3-9", "D16", "存货合计", "审计报告.pdf", 5, "t1", "7,5",
        "源值直取+单位换算", "资产负债表账面净额", "已填",
    ]]

    _append_mapping(
        mappings, sheet, cell, "存货合计", [],
        "附注合计直取", "存货附注账面余额",
    )

    assert mappings[0][-1] == "已被更精确证据替代"
    assert mappings[1][-1] == "已填"
    assert sum(row[-1] == "已填" for row in mappings) == 1
