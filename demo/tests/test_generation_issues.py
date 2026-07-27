from openpyxl import load_workbook

from demo.adapters.generation_issues import export_generation_issues
from demo.domain.generation_issues import issues_for_missing_locations
from demo.domain.replacement import build_replacements


def test_missing_field_uses_original_marker_and_creates_location_issue():
    location = {
        "location_id": "DOCUMENT-P0001-X01",
        "record_type": "占位符",
        "field_key": "company_name",
        "field_name": "公司名称",
        "marker": "XXX",
        "context": "公司名称：XXX",
        "source_kind": "人工输入",
        "source_file": "人工基础信息",
        "source_locator": "公司名称",
    }

    replacements = build_replacements([location], {})
    issues = issues_for_missing_locations([location], {})

    assert replacements[location["location_id"]] == "XXX"
    assert issues[0]["location_id"] == location["location_id"]
    assert issues[0]["current_text"] == "XXX"
    assert issues[0]["expected_source"] == "人工输入"


def test_missing_20xx_field_preserves_full_marker():
    location = {
        "location_id": "DOCUMENT-P0001-X01",
        "field_key": "valuation_year",
        "field_name": "评估年份",
        "marker": "20XX",
        "context": "评估基准日：20XX年",
    }

    assert build_replacements([location], {})[location["location_id"]] == "20XX"


def test_generation_issue_export_has_required_columns(tmp_path):
    output = export_generation_issues(
        tmp_path / "生成问题清单.xlsx",
        [
            {
                "issue_id": "GEN-1",
                "priority": "高",
                "category": "missing_field",
                "page_number": 9,
                "page_basis": "generated_report",
                "location_id": "DOCUMENT-P0100-X01",
                "location_type": "段落",
                "location_description": "公司名称：XXX",
                "field_key": "company_name",
                "field_name": "公司名称",
                "current_text": "XXX",
                "problem": "指定来源未匹配到可用值",
                "expected_source": "人工输入",
                "source_file": "",
                "source_locator": "公司名称",
                "suggestion": "补充公司名称",
                "status": "待人工处理",
            }
        ],
    )

    sheet = load_workbook(output, read_only=True)["生成问题"]
    assert sheet["D2"].value == 9
    assert sheet["F2"].value == "DOCUMENT-P0100-X01"
