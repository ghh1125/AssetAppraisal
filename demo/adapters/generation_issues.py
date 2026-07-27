from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


COLUMNS = [
    ("issue_id", "问题编号"),
    ("priority", "优先级"),
    ("category", "问题类别"),
    ("page_number", "Word页码"),
    ("page_basis", "页码依据"),
    ("location_id", "位置编号"),
    ("location_type", "位置类型"),
    ("location_description", "位置描述"),
    ("field_key", "标准字段"),
    ("field_name", "字段名称"),
    ("current_text", "当前内容"),
    ("problem", "问题说明"),
    ("expected_source", "应取来源"),
    ("source_file", "来源文件"),
    ("source_locator", "来源位置"),
    ("suggestion", "处理建议"),
    ("status", "处理状态"),
]


def export_generation_issues(
    output: Path,
    issues: list[dict],
) -> Path:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "生成问题"
    sheet.append([label for _, label in COLUMNS])
    for issue in issues:
        sheet.append([issue.get(key, "") for key, _ in COLUMNS])
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="FFF2CC")
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )
    widths = [16, 9, 18, 10, 18, 28, 14, 48, 28, 24, 22, 42, 22, 32, 42, 38, 14]
    for index, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output)
    return output
