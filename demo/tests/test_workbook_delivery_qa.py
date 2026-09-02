from pathlib import Path
from zipfile import ZipFile

import pytest

from demo.adapters.workbook_delivery_qa import prepare_workbook_for_delivery


def _write_minimal_workbook(path: Path, formula: str, cached: str) -> None:
    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheets><sheet name="Sheet1" sheetId="1"/></sheets>'
        '<calcPr calcId="1" calcMode="manual"/></workbook>'
    )
    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetData><row r="1"><c r="A1" t="e"><f>{formula}</f><v>{cached}</v></c></row></sheetData>'
        '</worksheet>'
    )
    with ZipFile(path, "w") as workbook:
        workbook.writestr("xl/workbook.xml", workbook_xml)
        workbook.writestr("xl/worksheets/sheet1.xml", sheet_xml)


def test_backend_clears_stale_cached_formula_errors_and_forces_recalculation(tmp_path):
    path = tmp_path / "model.xlsx"
    _write_minimal_workbook(path, "Sheet1!B1", "#NAME?")

    result = prepare_workbook_for_delivery(path)

    assert result["cached_formula_errors_cleared"] == 1
    with ZipFile(path) as workbook:
        workbook_xml = workbook.read("xl/workbook.xml").decode("utf-8")
        sheet_xml = workbook.read("xl/worksheets/sheet1.xml").decode("utf-8")
    assert 'calcMode="auto"' in workbook_xml
    assert 'fullCalcOnLoad="1"' in workbook_xml
    assert "#NAME?" not in sheet_xml
    assert '<f>Sheet1!B1</f>' in sheet_xml


def test_backend_rejects_structural_formula_errors(tmp_path):
    path = tmp_path / "broken.xlsx"
    _write_minimal_workbook(path, "#REF!+1", "#REF!")

    with pytest.raises(RuntimeError, match="结构性错误"):
        prepare_workbook_for_delivery(path)
