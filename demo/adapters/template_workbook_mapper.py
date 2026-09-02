"""Populate the two real appraisal-workbook templates from audit OCR evidence.

The templates are examples of workbook *structure*, not sources of financial
facts.  This adapter copies their layout/formulas, removes client-specific
inputs, fills only audit-proven historical/book amounts and records every
decision in an ``AI映射规则`` worksheet.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
import shutil
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .audit_intake import ExtractedTable, _ocr_number, tables_from_pages
from .note_evidence_mapper import enhance_from_notes
from .rule_graph_builder import build_rule_graph_bundle


@dataclass(frozen=True)
class SourceRow:
    table: ExtractedTable
    row: int
    label: str
    values: tuple[tuple[int, float], ...]
    periods: tuple[tuple[int, str], ...]
    unit_to_yuan: float | None
    scope: str


def _key(value: Any) -> str:
    text = str(value or "").replace("\n", "").strip()
    text = re.sub(r"^\s*\d+[、.．]\s*", "", text)
    text = re.sub(r"^(其中|[一二三四五六七八九十]+、|减：|加：)\s*", "", text)
    text = re.sub(r"[（(].*?[）)]", "", text)
    text = re.sub(r"(净额|合计|总计|账面价值|评估价值)$", "", text)
    text = re.sub(r"[\s：:—\-]", "", text)
    aliases = {
        "所有者权益": "所有者权益",
        "股东权益": "所有者权益",
        "应收账款净额": "应收账款",
        "应收具": "应收票据",
        "应收欧项融资": "应收款项融资",
        "预付账款净额": "预付款项",
        "预付账款": "预付款项",
        "预付欧项": "预付款项",
        "预收账款": "预收款项",
        "工程物质": "工程物资",
        "在则工程": "在建工程",
        "流动资产合计十": "流动资产",
        "长期应收款人": "长期应收款",
        "天开开发支出": "开发支出",
        "固定资产净额": "固定资产",
        "无形资产净额": "无形资产",
        "长期股权投资净额": "长期股权投资",
        "股本": "实收资本",
        "米分配利润": "未分配利润",
        "取得信款收到的现金": "取得借款收到的现金",
        "租货负债": "租赁负债",
        "净资产": "所有者权益",
        "应付账款": "应付账款",
        "营业收入": "营业收入",
        "营业总收入": "营业总收入",
        "营业成本": "营业成本",
        "净利润": "净利润",
        "资产总计": "资产",
        "负债合计": "负债",
        "负债和所有者权益总计": "负债和所有者权益",
        "所有者权益合计": "所有者权益",
        "股东权益合计": "所有者权益",
        "流动资产合计": "流动资产",
        "流动负债合计": "流动负债",
        "非流动资产合计": "非流动资产",
        "非流动负债合计": "非流动负债",
    }
    return aliases.get(text, text)


def _page_unit_to_yuan(pages: list[dict[str, Any]]) -> dict[int, float | None]:
    units: dict[int, float | None] = {}
    for page in pages:
        text = " ".join([
            *(str(block.get("text", "")) for block in page.get("blocks", [])),
            *(
                str(cell.get("text", ""))
                for table in page.get("tables", [])
                for cell in table.get("cells", [])
            ),
        ])
        if re.search(r"(?:金额)?单位(?:均为)?\s*[：:]?\s*(?:人民币)?\s*万元|人民币\s*万元", text):
            units[int(page.get("page_number") or 0)] = 10_000.0
        elif re.search(r"(?:金额)?单位(?:均为)?\s*[：:]?\s*(?:人民币)?\s*元|人民币\s*元", text):
            units[int(page.get("page_number") or 0)] = 1.0
        else:
            units[int(page.get("page_number") or 0)] = None
    # Inherit a document-wide unit only when every explicitly stated page uses
    # the same unit.  An absent or conflicting unit is never guessed.
    known = {value for value in units.values() if value is not None}
    if len(known) == 1:
        inherited = next(iter(known))
        units = {page: value if value is not None else inherited for page, value in units.items()}
    return units


def _table_periods(table: ExtractedTable) -> dict[int, str]:
    periods = dict(table.period_headers)
    header_rows = table.matrix[:3]
    max_column = max((len(row) for row in header_rows), default=0)
    header_text: dict[int, str] = {}
    for column in range(1, max_column + 1):
        parts = [str(row[column - 1] or "").strip() for row in header_rows if len(row) >= column]
        text = " ".join(parts).strip()
        header_text[column] = text
        interval = re.search(r"((?:19|20)\d{2})\s*[年.]?\s*(\d{1,2})\s*月?\s*[-至~]\s*(\d{1,2})\s*月?", text)
        if interval:
            year, _start_month, end_month = (int(value) for value in interval.groups())
            periods[column] = f"{year:04d}-{end_month:02d}-末"
            continue
        date_range = re.search(
            r"((?:19|20)\d{2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日\s*至\s*"
            r"((?:19|20)\d{2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日",
            text,
        )
        if date_range:
            sy, sm, sd, ey, em, ed = (int(value) for value in date_range.groups())
            periods[column] = f"{sy:04d}-{sm:02d}-{sd:02d}至{ey:04d}-{em:02d}-{ed:02d}"
            continue
        full_date = re.search(r"((?:19|20)\d{2})\s*[年./-]\s*(\d{1,2})\s*[月./-]\s*(\d{1,2})\s*日", text)
        if full_date:
            year, month, day = (int(value) for value in full_date.groups())
            periods[column] = f"{year:04d}-{month:02d}-{day:02d}"
            continue
        annual = next((re.fullmatch(r"\s*((?:19|20)\d{2})\s*年度\s*", part) for part in parts if re.fullmatch(r"\s*((?:19|20)\d{2})\s*年度\s*", part)), None)
        if annual:
            periods[column] = f"{int(annual.group(1)):04d}年度"
            continue
        year_only = next((re.fullmatch(r"\s*((?:19|20)\d{2})\s*年?\s*", part) for part in parts if re.fullmatch(r"\s*((?:19|20)\d{2})\s*年?\s*", part)), None)
        if year_only:
            periods[column] = f"{int(year_only.group(1)):04d}-12-31"
    # Merged period headers frequently span separate "公司/合并" columns.  The
    # date is present only in the left cell while the adjacent right cell has
    # a scope label on the next header row.
    first_header = {
        column: str(header_rows[0][column - 1] or "").strip() if header_rows and len(header_rows[0]) >= column else ""
        for column in range(1, max_column + 1)
    }
    for column in range(2, max_column + 1):
        if column not in periods and not first_header.get(column) and column - 1 in periods:
            periods[column] = periods[column - 1]
    return periods


def _table_column_scopes(table: ExtractedTable) -> dict[int, str]:
    result: dict[int, str] = {}
    header_rows = table.matrix[:3]
    max_column = max((len(row) for row in header_rows), default=0)
    for column in range(1, max_column + 1):
        parts = [str(row[column - 1] or "").strip() for row in header_rows if len(row) >= column]
        text = " ".join(parts)
        if "合并" in text or "合井" in text:
            result[column] = "合并"
        elif "母公司" in text or any(re.fullmatch(r"\s*[：:]?公司\s*", part) for part in parts):
            result[column] = "母公司"
    return result


def _page_scopes(pages: list[dict[str, Any]]) -> dict[int, str]:
    result: dict[int, str] = {}
    for page in pages:
        text = " ".join(str(block.get("text", "")) for block in page.get("blocks", []))
        if "母公司" in text or "单体" in text:
            result[int(page.get("page_number") or 0)] = "母公司"
        elif "合并" in text:
            result[int(page.get("page_number") or 0)] = "合并"
        else:
            result[int(page.get("page_number") or 0)] = "未识别"
    return result


def _relative_statement_periods(table: ExtractedTable, balance_date: str) -> dict[int, str]:
    """Resolve 本期/上期 columns from the audited balance-sheet date."""
    match = re.fullmatch(r"((?:19|20)\d{2})-(\d{2})-(\d{2})", balance_date)
    if not match:
        return {}
    year, month, day = (int(value) for value in match.groups())
    current = f"{year}年度" if month == 12 and day >= 28 else f"{year:04d}-{month:02d}-末"
    prior = f"{year - 1}年度" if month == 12 and day >= 28 else f"{year - 1:04d}-{month:02d}-末"
    result: dict[int, str] = {}
    header_rows = table.matrix[:2]
    max_column = max((len(row) for row in header_rows), default=0)
    for column in range(1, max_column + 1):
        text = " ".join(str(row[column - 1] or "") for row in header_rows if len(row) >= column)
        if re.search(r"本期|本年|本期累计|本年累计", text):
            result[column] = current
        elif re.search(r"上期|上年|上期累计|上年同期", text):
            result[column] = prior
    return result


def _source_rows(document_name: str, pages: list[dict[str, Any]]) -> dict[str, list[SourceRow]]:
    units = _page_unit_to_yuan(pages)
    scopes = _page_scopes(pages)
    result: dict[str, list[SourceRow]] = {}
    statement_tables = [
        table
        for table in tables_from_pages(document_name, pages)
        if table.category in {"资产负债表", "利润表", "现金流量表"}
    ]
    balance_dates = [
        period
        for table in statement_tables
        if table.category == "资产负债表"
        for period in _table_periods(table).values()
        if re.fullmatch(r"(?:19|20)\d{2}-\d{2}-\d{2}", period)
    ]
    balance_date = max(balance_dates, key=_period_sort_key) if balance_dates else ""
    for table in statement_tables:
        if table.category not in {"资产负债表", "利润表", "现金流量表"}:
            continue
        periods = _table_periods(table)
        if table.category in {"利润表", "现金流量表"}:
            for column, period in _relative_statement_periods(table, balance_date).items():
                periods.setdefault(column, period)
        column_scopes = _table_column_scopes(table)
        # A number of statutory statements place assets and liabilities (or
        # profit and appropriation) side by side in one OCR table.  Treating
        # the leftmost text as the only row label loses an entire half of the
        # statement.  We split each row into label-led horizontal segments and
        # exclude columns headed "行次" from monetary candidates.
        line_number_columns = {
            column
            for column in range(1, max((len(row) for row in table.matrix), default=0) + 1)
            if "行次" in "".join(str(row[column - 1] or "") for row in table.matrix[:3] if len(row) >= column)
        }
        annotation_columns = {
            column
            for column in range(1, max((len(row) for row in table.matrix), default=0) + 1)
            if re.search(
                r"附注|注释|注释号|附注号",
                "".join(str(row[column - 1] or "") for row in table.matrix[:3] if len(row) >= column),
            )
        }
        for row_number, row in enumerate(table.matrix, 1):
            labels = [
                (column, str(raw))
                for column, raw in enumerate(row, 1)
                if isinstance(raw, str)
                and raw.strip()
                # A statutory statement commonly has a textual note-reference
                # column between the account label and the monetary columns.
                # Treating values such as ``五(二十七)`` as a second row label
                # closes the account segment before any amount is reached and
                # silently drops the entire statement row.
                and column not in annotation_columns
                and not isinstance(_ocr_number(raw), (int, float))
                and _key(raw) not in {"", "项目", "行次"}
            ]
            for position, (start_column, label) in enumerate(labels):
                end_column = labels[position + 1][0] if position + 1 < len(labels) else len(row) + 1
                parsed_values = tuple(
                    (column, float(parsed))
                    for column in range(start_column + 1, end_column)
                    if column not in line_number_columns
                    and column in periods
                    and isinstance((parsed := _ocr_number(row[column - 1])), (int, float))
                )
                label_key = _key(label)
                grouped: dict[str, list[tuple[int, float]]] = {}
                for column, value in parsed_values:
                    scope = column_scopes.get(column, scopes.get(table.page_number, "未识别"))
                    grouped.setdefault(scope, []).append((column, value))
                if label_key:
                    for scope, scoped_values in grouped.items():
                        values = tuple(scoped_values)
                        result.setdefault(label_key, []).append(
                            SourceRow(
                                table,
                                row_number,
                                label,
                                values,
                                tuple((column, periods[column]) for column, _ in values if column in periods),
                                units.get(table.page_number),
                                scope,
                            )
                        )
    return result


def _find_source(label: Any, source_rows: dict[str, list[SourceRow]], preferred_scope: str | None = None) -> SourceRow | None:
    target = _key(label)
    if not target:
        return None
    exact = source_rows.get(target)
    if not exact:
        return None
    # Parent and consolidated statements can contain the same account name.
    # Without a project-level scope instruction, differing scopes are
    # ambiguous and must remain blank.  Identical repetitions are safe.
    if preferred_scope:
        preferred = [item for item in exact if item.scope == preferred_scope]
        # Once the case has selected a reporting scope, never fall back to a
        # different company's/consolidated column merely because the chosen
        # scope is blank for that account.  A blank parent-company amount is
        # evidence of absence, not permission to import the group figure.
        if not preferred:
            return None
        exact = preferred
    parent = [item for item in exact if item.scope == "母公司"]
    consolidated = [item for item in exact if item.scope == "合并"]
    unknown = [item for item in exact if item.scope == "未识别"]
    if parent and consolidated:
        scope_signatures = {
            tuple(round(value * (item.unit_to_yuan or 1.0), 6) for _, value in item.values)
            for item in [*parent, *consolidated]
        }
        if len(scope_signatures) != 1:
            return None
        candidates = [*parent, *consolidated]
    else:
        candidates = parent or consolidated or unknown
    # Financial-statement rows can show both original cost and carrying/net
    # amount.  Asset workpapers use the statement carrying amount; never pick
    # a cost row simply because its simplified account name is identical.
    net_amounts = [item for item in candidates if "净额" in item.label]
    if net_amounts:
        candidates = net_amounts
    signatures = {
        tuple(round(value * (item.unit_to_yuan or 1.0), 6) for _, value in item.values)
        for item in candidates
    }
    return max(candidates, key=lambda item: len(item.values)) if len(signatures) == 1 else None


def _row_label(sheet, row: int) -> str | None:
    candidates = [
        str(sheet.cell(row, column).value).strip()
        for column in range(1, min(sheet.max_column, 6) + 1)
        if isinstance(sheet.cell(row, column).value, str) and sheet.cell(row, column).value.strip()
    ]
    candidates = [item for item in candidates if not re.fullmatch(r"\d+(?:-\d+)+", item)]
    return max(candidates, key=len) if candidates else None


def _unit_of_sheet(sheet) -> float:
    text = " ".join(str(sheet.cell(row, column).value or "") for row in range(1, min(sheet.max_row, 7) + 1) for column in range(1, min(sheet.max_column, 16) + 1))
    return 10_000.0 if "万元" in text else 1.0


def _source_item(source: SourceRow, ordinal: int) -> tuple[int, float] | None:
    if source.unit_to_yuan is None:
        return None
    period_by_column = dict(source.periods)
    # Even a single numeric column is unsafe without a dated/period-labelled
    # header: it could be an opening balance, closing balance or note amount.
    if not source.values or not all(column in period_by_column for column, _ in source.values):
        return None
    ordered = sorted(
        source.values,
        key=lambda item: period_by_column.get(item[0], "0000-00-00"),
        reverse=True,
    )
    # OCR can split one period across two adjacent columns.  Keep only the
    # non-empty value already present in this row for that period.
    deduplicated: list[tuple[int, float]] = []
    seen_periods: set[str] = set()
    for item in ordered:
        period = period_by_column.get(item[0], f"column-{item[0]}")
        if period in seen_periods:
            continue
        seen_periods.add(period)
        deduplicated.append(item)
    if ordinal >= len(deduplicated):
        return None
    return deduplicated[ordinal]


def _source_value(source: SourceRow, ordinal: int, target_unit_to_yuan: float) -> float | None:
    item = _source_item(source, ordinal)
    if item is None or source.unit_to_yuan is None:
        return None
    return item[1] * source.unit_to_yuan / target_unit_to_yuan


def _period_sort_key(period: str) -> tuple[int, int, int]:
    range_end = re.search(r"至((?:19|20)\d{2})-(\d{2})-(\d{2})$", period)
    if range_end:
        return tuple(int(value) for value in range_end.groups())
    match = re.match(r"((?:19|20)\d{2})(?:-(\d{1,2}))?(?:-(\d{1,2}))?", period)
    if not match:
        return 0, 0, 0
    year = int(match.group(1)); month = int(match.group(2) or 12); day = int(match.group(3) or 31)
    return year, month, day


def _normalize_period(category: str, period: str) -> str:
    """Put statement-specific period labels on one comparable axis."""
    year_end = re.fullmatch(r"((?:19|20)\d{2})-12-31", period)
    if category in {"利润表", "现金流量表"} and year_end:
        return f"{year_end.group(1)}年度"
    annual = re.fullmatch(r"((?:19|20)\d{2})年度", period)
    if category == "资产负债表" and annual:
        return f"{annual.group(1)}-12-31"
    return period


def _period_label(period: str) -> str:
    if period.endswith("年度"):
        return period
    partial = re.fullmatch(r"((?:19|20)\d{2})-(\d{2})-末", period)
    if partial:
        year, month = (int(value) for value in partial.groups())
        return f"{year}年1-{month}月"
    date_range = re.fullmatch(
        r"((?:19|20)\d{2})-(\d{2})-(\d{2})至((?:19|20)\d{2})-(\d{2})-(\d{2})",
        period,
    )
    if date_range:
        sy, sm, sd, ey, em, ed = (int(value) for value in date_range.groups())
        return f"{sy}年{sm}月{sd}日至{ey}年{em}月{ed}日"
    match = re.fullmatch(r"((?:19|20)\d{2})-(\d{2})-(\d{2})", period)
    if match:
        year, month, day = (int(value) for value in match.groups())
        return f"{year}年{month}月{day}日"
    return period


def _history_periods(source_rows: dict[str, list[SourceRow]], category: str, preferred_scope: str) -> list[str]:
    counts: dict[str, int] = {}
    items = [item for candidates in source_rows.values() for item in candidates if item.table.category == category]
    preferred = [item for item in items if item.scope == preferred_scope]
    items = preferred or items
    for item in items:
        value_columns = {column for column, _value in item.values}
        for column, period in item.periods:
            if column in value_columns:
                period = _normalize_period(category, period)
                counts[period] = counts.get(period, 0) + 1
    if not counts:
        return []
    # A real statement period appears throughout the body.  Dates occurring
    # only in a signature/footer line must not become financial periods.
    threshold = max(1, (max(counts.values()) + 1) // 2)
    eligible = [period for period, count in counts.items() if count >= threshold]
    return sorted(eligible, key=_period_sort_key, reverse=True)[:4]


def _source_value_for_period(source: SourceRow, period: str, target_unit_to_yuan: float) -> tuple[int, float] | None:
    if source.unit_to_yuan is None:
        return None
    period_by_column = dict(source.periods)
    matches = [
        (column, value)
        for column, value in source.values
        if _normalize_period(source.table.category, period_by_column.get(column, "")) == period
    ]
    if not matches:
        return None
    values = {round(value, 9) for _column, value in matches}
    if len(values) != 1:
        return None
    column, value = matches[0]
    return column, value * source.unit_to_yuan / target_unit_to_yuan


def _add_mapping_sheet(workbook, rows: list[list[Any]]) -> None:
    if "AI映射规则" in workbook.sheetnames:
        del workbook["AI映射规则"]
    sheet = workbook.create_sheet("AI映射规则")
    sheet.append(["目标工作表", "目标单元格", "模板字段", "来源文件", "来源页", "来源表", "来源行", "取数方式", "计算/处理", "状态"])
    for row in rows:
        sheet.append(row)
    for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:J{sheet.max_row}"
    for column in "ABCDEFGHIJ":
        sheet.column_dimensions[column].width = 18


def _add_formula_rules_sheet(workbook) -> None:
    """Expose every retained template formula as a cell-level rule."""
    if "AI公式规则" in workbook.sheetnames:
        del workbook["AI公式规则"]
    sheet = workbook.create_sheet("AI公式规则")
    sheet.append(["目标工作表", "目标单元格", "公式", "规则类型", "说明"])
    for source_sheet in workbook.worksheets:
        if source_sheet.title in {"AI映射规则", "AI公式规则"}:
            continue
        for row in source_sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    sheet.append([source_sheet.title, cell.coordinate, cell.value, "模板公式", "由 Excel 按公式引用的单元格自动计算"])
    for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.font = Font(color="FFFFFF", bold=True)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:E{sheet.max_row}"
    sheet.column_dimensions["A"].width = 24
    sheet.column_dimensions["B"].width = 14
    sheet.column_dimensions["C"].width = 42
    sheet.column_dimensions["D"].width = 16
    sheet.column_dimensions["E"].width = 36


def _add_combined_mapping_sheet(workbook, asset_rows: list[list[Any]], income_rows: list[list[Any]]) -> None:
    """Write the mapping ledger outside the two protected business templates."""
    sheet = workbook.create_sheet("AI映射规则")
    headers = ["目标文件", "目标工作表", "目标单元格", "模板字段", "来源文件", "来源页", "来源表", "来源行", "取数方式", "计算/处理", "状态"]
    sheet.append(headers)
    for target_file, rows in (("资产基础法_资产清查.xlsx", asset_rows), ("收益法_市场法.xlsx", income_rows)):
        for row in rows:
            sheet.append([target_file, *row])
    for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:K{sheet.max_row}"
    for column in "ABCDEFGHIJK":
        sheet.column_dimensions[column].width = 18


def _add_formula_rules_from(workbook, named_sources: list[tuple[str, Any]]) -> None:
    sheet = workbook.create_sheet("AI公式规则")
    sheet.append(["目标文件", "目标工作表", "目标单元格", "公式", "规则类型", "说明"])
    for target_file, source_workbook in named_sources:
        for source_sheet in source_workbook.worksheets:
            for row in source_sheet.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        sheet.append([target_file, source_sheet.title, cell.coordinate, cell.value, "模板公式", "由 Excel 按公式引用的单元格自动计算"])
    for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.font = Font(color="FFFFFF", bold=True)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:F{sheet.max_row}"
    sheet.column_dimensions["A"].width = 26
    sheet.column_dimensions["B"].width = 24
    sheet.column_dimensions["C"].width = 14
    sheet.column_dimensions["D"].width = 42
    sheet.column_dimensions["E"].width = 16
    sheet.column_dimensions["F"].width = 36


def _required_material(target_file: str, sheet_name: str, field: str) -> str:
    text = f"{sheet_name} {field}"
    if target_file.startswith("资产基础法"):
        if any(token in text for token in ("评估价值", "评估原值", "评估净值", "成新率", "重置", "增减", "增值")):
            return "评估作价依据、询价记录或评估参数"
        if any(token in text for token in ("设备", "规格", "生产厂家", "数量", "购置日期", "启用日期")):
            return "设备清单、固定资产卡片及现场盘点资料"
        if any(token in text for token in ("欠款", "户名", "结算对象", "账龄", "发生日期")):
            return "科目余额表及往来明细/账龄表"
        if any(token in text for token in ("被投资单位", "持股", "投资成本")):
            return "长期股权投资明细、投资协议及被投资单位资料"
        if "账面" in text or "帐面" in text:
            return "审计报表、附注或科目余额表"
        return "对应资产清查明细及权属证明"
    if "折现率" in text:
        return "无风险利率、Beta、资本结构、债务成本及市场参数"
    if any(token in text for token in ("收入", "成本", "费用", "收益")):
        return "管理层预测、历史分部数据及预测依据"
    if any(token in text for token in ("折旧", "摊销", "资本性支出")):
        return "固定资产明细、折旧政策及资本性支出计划"
    if "营运资金" in text:
        return "历史营运资金明细及预测周转假设"
    if "所得税" in text or "增值税" in text or "税金" in text:
        return "适用税率、纳税申报及税收优惠文件"
    return "管理层预测或对应专项明细"


def _add_template_field_registry(
    workbook,
    asset,
    income,
    asset_rows: list[list[Any]],
    income_rows: list[list[Any]],
    income_value_bounds: dict[str, tuple[int, int]],
) -> None:
    sheet = workbook.create_sheet("模板字段注册表")
    sheet.append(["目标文件", "目标工作表", "目标单元格", "字段/列", "字段类型", "当前状态", "来源/处理说明", "缺失时需要补充资料"])
    mapping_index: dict[tuple[str, str, str], list[Any]] = {}
    for target_file, rows in (("资产基础法_资产清查.xlsx", asset_rows), ("收益法_市场法.xlsx", income_rows)):
        for row in rows:
            target_cell = str(row[1] or "")
            if ":" not in target_cell:
                mapping_index[(target_file, str(row[0]), target_cell)] = row

    def append(target_file: str, source_sheet, cell, field: str, field_type: str) -> None:
        mapping = mapping_index.get((target_file, source_sheet.title, cell.coordinate))
        if isinstance(cell.value, str) and cell.value.startswith("="):
            status, note = "模板公式", cell.value
        elif mapping:
            status, note = str(mapping[-1]), str(mapping[-2])
        elif cell.value not in (None, ""):
            status, note = "模板结构值", "保留模板结构常量；不作为本项目事实来源"
        else:
            status, note = "待补充", "材料未提供可验证依据"
        sheet.append([target_file, source_sheet.title, cell.coordinate, field, field_type, status, note, _required_material(target_file, source_sheet.title, field) if status in {"未填", "待补充"} else ""])

    for source_sheet in asset.worksheets:
        if not source_sheet.title.startswith("表") or source_sheet.max_row < 8:
            continue
        headers = {column: str(source_sheet.cell(7, column).value or "") for column in range(1, min(source_sheet.max_column, 24) + 1)}
        detail = "-" in source_sheet.title and source_sheet.max_row >= 48
        for column, header in headers.items():
            if not header:
                continue
            if detail:
                if header in {"序号"}:
                    continue
                row_range = range(8, 48)
                field_type = "明细输入/计算"
            else:
                if not any(token in header for token in ("账面", "帐面", "评估", "增减", "增值", "原值", "净值")):
                    continue
                row_range = range(8, source_sheet.max_row + 1)
                field_type = "汇总输入/结果"
            for row in row_range:
                append("资产基础法_资产清查.xlsx", source_sheet, source_sheet.cell(row, column), header, field_type)

    for name in ("历资表", "历利表", "历现表"):
        source_sheet = income[name]
        for row in range(5, source_sheet.max_row + 1):
            label = str(source_sheet.cell(row, 1).value or "")
            if not label:
                continue
            for column in range(2, 6):
                append("收益法_市场法.xlsx", source_sheet, source_sheet.cell(row, column), label, "历史财务输入")

    input_style_ids = {5, 12, 23, 42, 43, 44, 63, 134, 135, 138, 139, 140, 141}
    for source_sheet in income.worksheets:
        if source_sheet.title in {"目录1", "目录2", "项目信息", "历资表", "历利表", "历现表"}:
            continue
        max_row, max_column = income_value_bounds.get(source_sheet.title, (0, 0))
        if max_row < 4 or max_column < 1:
            continue
        for row in source_sheet.iter_rows(min_row=4, max_row=max_row, max_col=max_column):
            for cell in row:
                fill = cell.fill.fgColor.rgb if cell.fill and cell.fill.fill_type else None
                if cell.style_id not in input_style_ids and fill not in {"FFFEFEFE", "FF90EE90"}:
                    continue
                append("收益法_市场法.xlsx", source_sheet, cell, str(source_sheet.cell(cell.row, 1).value or source_sheet.title), "预测/假设输入")

    for target_cell, field in (("B5", "企业名称"), ("B7", "评估基准日"), ("B8", "金额单位")):
        append("收益法_市场法.xlsx", income["项目信息"], income["项目信息"][target_cell], field, "项目信息")

    for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:H{sheet.max_row}"
    widths = [28, 25, 14, 28, 18, 16, 45, 45]
    for index, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(index)].width = width


def _add_note_candidates(workbook, document_name: str, pages: list[dict[str, Any]]) -> None:
    sheet = workbook.create_sheet("附注及明细候选")
    sheet.append(["来源文件", "页码", "OCR表格", "分类", "表头/内容预览", "自动处理状态", "原因"])
    for table in tables_from_pages(document_name, pages):
        if table.category not in {"资产明细", "未分类财务表", "所有者权益变动表"}:
            continue
        preview = " | ".join(str(value or "") for row in table.matrix[:3] for value in row if str(value or "").strip())[:500]
        sheet.append([document_name, table.page_number, table.table_id, table.category, preview, "未自动写入", "需与目标明细列、期间、单位及记录主键同时匹配；当前仅保留为可追溯候选"])
    for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.font = Font(color="FFFFFF", bold=True)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:G{sheet.max_row}"
    sheet.column_dimensions["A"].width = 36
    sheet.column_dimensions["E"].width = 80
    sheet.column_dimensions["G"].width = 56


def _add_audit_source_sheets(workbook, document_name: str, pages: list[dict[str, Any]]) -> None:
    """Keep every OCR table in the generated workbook, not just matched rows."""
    for sheet in list(workbook.worksheets):
        if sheet.title.startswith("审计原始_") or sheet.title == "审计原始索引":
            del workbook[sheet.title]
    index = workbook.create_sheet("审计原始索引")
    index.append(["来源文件", "PDF页码", "报表口径", "OCR表格编号", "表格类型", "输出工作表", "页内标题"])
    scopes = _page_scopes(pages)
    block_text = {
        int(page.get("page_number") or 0): " ".join(str(block.get("text", "")) for block in page.get("blocks", []))[:500]
        for page in pages
    }
    used = set(workbook.sheetnames)
    for number, table in enumerate(tables_from_pages(document_name, pages), 1):
        base = re.sub(r"[\\/*?:\[\]]", "_", f"审计原始_{table.page_number}_{table.category}_{number}")[:31]
        title = base
        suffix = 2
        while title in used:
            title = f"{base[:28]}_{suffix}"
            suffix += 1
        used.add(title)
        sheet = workbook.create_sheet(title)
        sheet.append([f"来源：{document_name}；第{table.page_number}页；{table.category}；OCR表：{table.table_id}"])
        for row in table.matrix:
            sheet.append([_ocr_number(value) for value in row])
        sheet.freeze_panes = "A2"
        for column in range(1, min(sheet.max_column, 20) + 1):
            sheet.column_dimensions[get_column_letter(column)].width = 18
        index.append([document_name, table.page_number, scopes.get(table.page_number, "未识别"), table.table_id, table.category, title, block_text.get(table.page_number, "")])
    for cell in index[1]:
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.font = Font(color="FFFFFF", bold=True)
    index.freeze_panes = "A2"
    index.auto_filter.ref = f"A1:G{index.max_row}"
    index.column_dimensions["A"].width = 32
    index.column_dimensions["G"].width = 72


def _add_supporting_material_sheets(workbook, materials: list[dict[str, Any]]) -> None:
    """Retain non-financial archive materials as traceable evidence.

    Business licences, company-information reports and screenshots are useful
    for project identity, but are never treated as financial-statement inputs.
    Keeping their OCR result in the same deliverable makes the boundary
    explicit and lets a reviewer trace every populated project field.
    """
    for sheet in list(workbook.worksheets):
        if sheet.title.startswith("企业资料_") or sheet.title == "企业资料索引":
            del workbook[sheet.title]
    index = workbook.create_sheet("企业资料索引")
    index.append(["来源文件", "材料类型", "OCR页码", "输出工作表", "识别主体", "统一社会信用代码", "处理原则"])
    used = set(workbook.sheetnames)
    sequence = 0
    for material in materials:
        pages = material.get("pages") or []
        metadata = material.get("metadata") or {}
        for page in pages:
            sequence += 1
            base = f"企业资料_{sequence}_{material.get('kind', '资料')}"
            title = re.sub(r"[\\/*?:\[\]]", "_", base)[:31]
            suffix = 2
            while title in used:
                title = f"{base[:28]}_{suffix}"
                suffix += 1
            used.add(title)
            sheet = workbook.create_sheet(title)
            page_number = int(page.get("page_number") or 0)
            sheet.append([f"来源：{material.get('document_name', '')}；第{page_number}页；{material.get('kind', '企业资料')}"])
            text = "\n".join(str(block.get("text") or "") for block in page.get("blocks", []))
            if text:
                sheet.append([text])
                sheet["A2"].alignment = Alignment(wrap_text=True, vertical="top")
                sheet.row_dimensions[2].height = 80
            for table in page.get("tables", []):
                sheet.append([f"OCR表：{table.get('table_id', '')}"])
                rows: dict[int, dict[int, str]] = {}
                for cell in table.get("cells", []):
                    rows.setdefault(int(cell.get("row") or 1), {})[int(cell.get("column") or 1)] = str(cell.get("text") or "")
                for row_number in sorted(rows):
                    values = rows[row_number]
                    sheet.append([values.get(column, "") for column in range(1, max(values, default=1) + 1)])
            sheet.column_dimensions["A"].width = 72
            index.append([
                material.get("document_name", ""), material.get("kind", "企业资料"), page_number,
                title, metadata.get("company_name", ""), metadata.get("credit_code", ""),
                material.get("match_status", "仅作为主体/项目资料证据；不参与财务金额取数"),
            ])
    for cell in index[1]:
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.font = Font(color="FFFFFF", bold=True)
    index.freeze_panes = "A2"
    index.auto_filter.ref = f"A1:G{index.max_row}"
    index.column_dimensions["A"].width = 36
    index.column_dimensions["E"].width = 30
    index.column_dimensions["G"].width = 42


def _write_metadata_sheet(workbook, metadata: dict[str, str], source_file: str) -> None:
    """Present verified project identity fields without guessing template cells."""
    if "项目主体信息" in workbook.sheetnames:
        del workbook["项目主体信息"]
    sheet = workbook.create_sheet("项目主体信息")
    sheet.append(["字段", "识别值", "来源", "取数方式"])
    labels = {
        "company_name": "企业名称",
        "credit_code": "统一社会信用代码",
        "legal_representative": "法定代表人",
        "address": "住所/注册地址",
        "business_scope": "经营范围",
        "registered_capital": "注册资本",
        "report_date": "审计报告期末日",
    }
    for key, label in labels.items():
        value = metadata.get(key, "")
        sheet.append([label, value, source_file if value else "", "OCR识别后待人工复核" if value else "未从材料可靠识别"])
    for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.font = Font(color="FFFFFF", bold=True)
    sheet.column_dimensions["A"].width = 24
    sheet.column_dimensions["B"].width = 54
    sheet.column_dimensions["C"].width = 34
    sheet.column_dimensions["D"].width = 30


def _basis_date(source_rows: dict[str, list[SourceRow]]) -> str:
    periods = _history_periods(source_rows, "资产负债表", "母公司")
    dated = [period for period in periods if re.fullmatch(r"(?:19|20)\d{2}-\d{2}-\d{2}", period)]
    return max(dated, key=_period_sort_key) if dated else ""


def _apply_project_identity(asset, income, metadata: dict[str, str], basis_date: str) -> None:
    company_name = metadata.get("company_name", "")
    subject = company_name or "待从审计报告/营业执照核实"
    legal_raw = str(metadata.get("legal_representative") or "").strip()
    legal_representative = re.split(r"[。；;，,\s]", legal_raw, maxsplit=1)[0]
    if not re.fullmatch(r"[\u4e00-\u9fff·]{2,8}", legal_representative):
        legal_representative = "待补充"
    date_match = re.fullmatch(r"((?:19|20)\d{2})-(\d{2})-(\d{2})", basis_date or "")
    chinese_date = ""
    if date_match:
        year, month, day = (int(value) for value in date_match.groups())
        chinese_date = f"{year}年 {month:02d}月 {day:02d}日"
    for sheet in asset.worksheets:
        for cell in sheet._cells.values():
            if isinstance(cell.value, str) and cell.value.startswith("="):
                continue
            if isinstance(cell.value, str):
                cell.value = cell.value.replace("通富热处理（昆山）有限公司", subject).replace("通富昆山", subject)
                if "评估机构" in cell.value:
                    cell.value = re.sub(r"评估机构\s*[：:].*", "评估机构: 待补充", cell.value)
                if "法定代表人" in cell.value or "项目负责人" in cell.value:
                    cell.value = f"法定代表人: {legal_representative}                                  项目负责人: 待补充"
                if isinstance(cell.value, str) and "被评估单位：" in cell.value:
                    cell.value = re.sub(r"被评估单位：.*", f"被评估单位：{subject}", cell.value)
                if isinstance(cell.value, str) and "评估基准日" in cell.value:
                    cell.value = re.sub(r"评估基准日\s*[：:].*", f"评估基准日:  {chinese_date or '待核实'}", cell.value)
                if isinstance(cell.value, str) and "填表日期" in cell.value:
                    cell.value = re.sub(r"填表日期\s*[：:].*", f"填表日期:  {chinese_date or '待填写'}", cell.value)
                if isinstance(cell.value, str) and chinese_date and re.fullmatch(r"\s*(?:19|20)\d{2}年\s*\d{1,2}月\s*\d{1,2}日\s*", cell.value):
                    cell.value = f" {chinese_date}"
    if company_name:
        income["项目信息"]["B5"] = company_name
    else:
        income["项目信息"]["B5"] = "未从审计报告/营业执照可靠识别"
    income["项目信息"]["B7"] = basis_date or "待按审计报告期末日期填写"
    # These cells are appraisal/model decisions in the Tongfu example, not
    # structural constants.  Keep the field positions but never inherit the
    # example company's perpetual-period, discounting or valuation choices.
    for row in (6, 9, 10, 11, 12, 13, 14, 15, 16, 17):
        income["项目信息"].cell(row, 2).value = "待补充评估参数"
    for sheet in income.worksheets:
        for row in sheet.iter_rows(max_row=min(sheet.max_row, 3), max_col=min(sheet.max_column, 20)):
            for cell in row:
                if isinstance(cell.value, str) and "评估基准日" in cell.value:
                    cell.value = re.sub(r"评估基准日\s*[：:].*", f"评估基准日：{basis_date or '待核实'}", cell.value)


def _apply_model_period_headers(income, actual_periods: list[str], basis_date: str) -> None:
    date_match = re.fullmatch(r"((?:19|20)\d{2})-(\d{2})-(\d{2})", basis_date or "")
    forecast_labels: list[str] = []
    if date_match:
        year, month, day = (int(value) for value in date_match.groups())
        if month == 12:
            forecast_labels = [f"{year + offset}年度" for offset in range(1, 7)] + ["永续年度"]
        else:
            forecast_labels = [f"{year}年{month + 1}-12月"] + [f"{year + offset}年度" for offset in range(1, 6)] + ["永续年度"]
    for sheet in income.worksheets:
        for row in range(1, min(sheet.max_row, 8) + 1):
            historical_column = next((column for column in range(1, min(sheet.max_column, 20) + 1) if sheet.cell(row, column).value == "历史数据"), None)
            forecast_column = next((column for column in range(1, min(sheet.max_column, 20) + 1) if sheet.cell(row, column).value == "未来预测"), None)
            if historical_column is None or forecast_column is None or forecast_column <= historical_column:
                continue
            header_row = row + 1
            history_columns = list(range(historical_column, forecast_column))
            for column in history_columns:
                sheet.cell(header_row, column).value = None
            for period, column in zip(actual_periods, reversed(history_columns)):
                sheet.cell(header_row, column).value = _period_label(period)
            for offset, column in enumerate(range(forecast_column, min(sheet.max_column, forecast_column + len(forecast_labels) - 1) + 1)):
                if offset < len(forecast_labels):
                    sheet.cell(header_row, column).value = forecast_labels[offset]
            break


def _set_recalculation(workbook) -> None:
    """Force Excel to calculate retained formulas when the file is opened."""
    calculation = getattr(workbook, "calculation", None)
    if calculation is not None:
        calculation.calcMode = "auto"
        calculation.fullCalcOnLoad = True
        calculation.forceFullCalc = True


def _write_rule_graph(
    output_dir: Path,
    document_name: str,
    asset_rows: list[list[Any]],
    income_rows: list[list[Any]],
    supporting_materials: list[dict[str, Any]],
) -> None:
    """Write a compact Mermaid graph beside the two generated workbooks."""
    lines = ["flowchart LR", f'  pdf["{document_name}"] --> index["审计原始索引"]']
    lines.append('  index --> financial["财务三表/附注 OCR 原表"]')
    lines.append('  financial --> asset_book["资产基础法：账面值"]')
    lines.append('  financial --> income_history["收益法：历史三表"]')
    for number, material in enumerate(supporting_materials, 1):
        material_id = f"support_{number}"
        label = str(material.get("document_name", "企业资料")).replace('"', "'")
        lines.append(f'  {material_id}["{label}"] --> support_index["企业资料索引"]')
    lines.extend([
        '  support_index --> identity["项目主体信息"]',
        '  identity -.主体匹配后.-> project["项目信息/表头"]',
        '  identity -.未匹配.-> review["保留原文，不写入本项目"]',
    ])
    seen: set[str] = set()
    for prefix, rows in (("asset", asset_rows), ("income", income_rows)):
        for row in rows:
            sheet, _, _, source_file, page, _, _, method, _, status = row
            if status != "已填" or not page:
                continue
            source_id = f"p{page}".replace("-", "_")
            target_id = re.sub(r"[^A-Za-z0-9_]", "_", f"{prefix}_{sheet}")
            edge = f'  {source_id}["PDF 第{page}页"] -->|{method}| {target_id}["{sheet}"]'
            if edge not in seen:
                lines.append(edge); seen.add(edge)
    lines.extend([
        '  index --> asset_raw["上报表：审计原始表"]',
        '  index --> income_raw["收益法：审计原始表"]',
        '  asset_raw --> asset_formula["表1/表2/明细汇总公式"]',
        '  income_raw --> income_formula["历史三表→预测/净现金流/WACC公式"]',
    ])
    (output_dir / "规则图.mmd").write_text("\n".join(lines) + "\n", encoding="utf-8")


def _clear_nonformula_numbers(workbook, excluded: set[str]) -> None:
    for sheet in workbook.worksheets:
        if sheet.title in excluded:
            continue
        for row in sheet.iter_rows(min_row=4):
            for cell in row:
                if isinstance(cell.value, (int, float)):
                    cell.value = None


def _clear_asset_sample_data(workbook) -> None:
    """Clear Tongfu inputs while preserving row labels, serials and formulas."""
    for sheet in workbook.worksheets:
        if not sheet.title.startswith("表") or sheet.max_row < 8:
            continue
        # Fixed 40-row detail schedules have a row-7 field header and a row-48
        # total.  Their records are entirely company-specific.
        if "-" in sheet.title and sheet.max_row >= 48:
            for row in sheet.iter_rows(min_row=8, max_row=47):
                for cell in row:
                    if not (isinstance(cell.value, str) and cell.value.startswith("=")):
                        cell.value = None
            continue
        # Summary schedules retain their structural sequence/code/account
        # labels.  Only numeric result columns are cleared.
        for column in range(1, min(sheet.max_column, 24) + 1):
            header = str(sheet.cell(7, column).value or "")
            if not any(token in header for token in ("账面", "帐面", "评估", "增减", "增值", "原值", "净值")):
                continue
            for row in range(8, sheet.max_row + 1):
                cell = sheet.cell(row, column)
                if not (isinstance(cell.value, str) and cell.value.startswith("=")):
                    cell.value = None


def _clear_income_sample_text(workbook) -> None:
    """Remove non-numeric Tongfu input labels left in the model."""
    input_fills = {"FFFEFEFE", "FF90EE90"}
    for sheet in workbook.worksheets:
        if sheet.title in {"目录1", "目录2", "项目信息"}:
            continue
        for row in sheet.iter_rows(min_row=4):
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    continue
                fill = cell.fill.fgColor.rgb if cell.fill and cell.fill.fill_type else None
                company_text = isinstance(cell.value, str) and any(token in cell.value for token in ("通富", "昆山", "热处理", "仙桃"))
                if company_text or (fill in input_fills and cell.value not in (None, "")):
                    cell.value = None


def _nonempty_bounds(workbook) -> dict[str, tuple[int, int]]:
    """Return content bounds without counting style-only whole-column cells."""
    result: dict[str, tuple[int, int]] = {}
    for sheet in workbook.worksheets:
        valued = [cell for cell in sheet._cells.values() if cell.value not in (None, "")]
        result[sheet.title] = (
            max((cell.row for cell in valued), default=0),
            max((cell.column for cell in valued), default=0),
        )
    return result


def _fill_history(
    workbook,
    sheet_name: str,
    category: str,
    source_rows: dict[str, list[SourceRow]],
    mappings: list[list[Any]],
    preferred_scope: str = "合并",
) -> list[str]:
    sheet = workbook[sheet_name]
    periods = _history_periods(source_rows, category, preferred_scope)
    period_targets = {period: column for period, column in zip(periods, (5, 4, 3, 2))}
    for column in range(2, 6):
        sheet.cell(4, column).value = None
    for period, column in period_targets.items():
        sheet.cell(4, column).value = _period_label(period)
    for row in range(5, sheet.max_row + 1):
        target_label = str(sheet.cell(row, 1).value or "")
        if not target_label.strip():
            continue
        # Section headings (e.g. ``流动资产：`` or ``经营活动产生的现金流量：``)
        # are presentation cells, not monetary fields.  Their normalized text
        # can collide with a later ``合计/净额`` source row, so exclude them
        # before semantic matching.
        if target_label.rstrip().endswith(("：", ":")):
            continue
        # In the general-format profit statement, row 10 is banking/insurance
        # interest revenue.  An ordinary company's later financial-expense
        # line ``利息收入`` must map to row 30, not be duplicated here.
        if sheet_name == "历利表" and row == 10:
            continue
        source = _find_source(target_label, source_rows, preferred_scope)
        if source is None:
            for period, target_column in period_targets.items():
                target = sheet.cell(row, target_column)
                if not (isinstance(target.value, str) and target.value.startswith("=")):
                    mappings.append([sheet.title, target.coordinate, sheet.cell(row, 1).value, "", "", "", "", "待补充", f"期间：{period}；未找到{preferred_scope}口径唯一同名标准科目", "未填"])
            continue
        for period, target_column in period_targets.items():
            source_value = _source_value_for_period(source, period, 10_000.0)
            if source_value is None:
                target = sheet.cell(row, target_column)
                if not (isinstance(target.value, str) and target.value.startswith("=")):
                    mappings.append([sheet.title, target.coordinate, sheet.cell(row, 1).value, source.table.source_file, source.table.page_number, source.table.table_id, source.row, "待补充", f"审计材料未提供{preferred_scope}口径的{period}金额", "未填"])
                continue
            target = sheet.cell(row, target_column)
            if isinstance(target.value, str) and target.value.startswith("="):
                continue
            source_column, value = source_value
            target.value = value
            source_unit = "万元" if source.unit_to_yuan == 10_000.0 else "元"
            mappings.append([sheet.title, target.coordinate, sheet.cell(row, 1).value, source.table.source_file, source.table.page_number, source.table.table_id, f"{source.row},{source_column}", "源值直取+单位换算", f"来源期间：{period}；来源单位：{source_unit}；审计金额÷10,000（模板单位：万元）；口径：{source.scope}", "已填"])
    return periods


def _fill_asset_template(workbook, source_rows: dict[str, list[SourceRow]], mappings: list[list[Any]], preferred_scope: str = "母公司") -> None:
    balance_periods = _history_periods(source_rows, "资产负债表", preferred_scope)
    basis_period = max(
        (period for period in balance_periods if re.fullmatch(r"(?:19|20)\d{2}-\d{2}-\d{2}", period)),
        key=_period_sort_key,
        default="",
    )
    for sheet in workbook.worksheets:
        if not sheet.title.startswith("表") or sheet.max_row < 8:
            continue
        header = {str(sheet.cell(7, col).value or ""): col for col in range(1, min(sheet.max_column, 20) + 1)}
        book_column = next((col for label, col in header.items() if "账面" in label or "帐面" in label), None)
        appraisal_column = next((col for label, col in header.items() if "评估价值" in label or "评估净值" in label), None)
        change_column = next((col for label, col in header.items() if "增减值" in label and "率" not in label), None)
        rate_column = next((col for label, col in header.items() if "增值率" in label), None)
        if book_column is None:
            continue
        target_unit = _unit_of_sheet(sheet)
        for row in range(8, sheet.max_row + 1):
            label = _row_label(sheet, row)
            # Do not create or replace client-facing formulas.  The original
            # Tongfu template is the sole authority for formula structure;
            # appraisal value/change columns are therefore cleared only when
            # they were sample constants, while retained formulas stay intact.
            if appraisal_column is not None and not (isinstance(sheet.cell(row, appraisal_column).value, str) and sheet.cell(row, appraisal_column).value.startswith("=")):
                sheet.cell(row, appraisal_column).value = None
            source = _find_source(label, source_rows, preferred_scope)
            if source is None:
                mappings.append([sheet.title, sheet.cell(row, book_column).coordinate, label, "", "", "", "", "待补充", "未找到唯一同名标准科目，或母公司/合并口径冲突", "未填"])
                continue
            source_value = _source_value_for_period(source, basis_period, target_unit) if basis_period else None
            if source_value is None:
                mappings.append([sheet.title, sheet.cell(row, book_column).coordinate, label, source.table.source_file, source.table.page_number, source.table.table_id, source.row, "待复核", f"来源未提供{preferred_scope}口径基准日{basis_period or '金额'}，不得以以前年度或其他口径替代", "未填"])
                continue
            source_column, value = source_value
            target = sheet.cell(row, book_column)
            target.value = value
            source_period = dict(source.periods).get(source_column, "")
            source_unit = "万元" if source.unit_to_yuan == 10_000.0 else "元"
            target_unit_label = "万元" if target_unit == 10_000.0 else "元"
            mappings.append([sheet.title, target.coordinate, label, source.table.source_file, source.table.page_number, source.table.table_id, f"{source.row},{source_column}", "源值直取+单位换算", f"来源期间：{source_period}；来源单位：{source_unit}；换算至模板单位：{target_unit_label}；口径：{source.scope}", "已填"])


def build_template_workbooks(
    *,
    output_dir: Path,
    document_name: str,
    pages: list[dict[str, Any]],
    asset_template: Path,
    income_template: Path,
    project_metadata: dict[str, str] | None = None,
    supporting_materials: list[dict[str, Any]] | None = None,
    progress_callback: Any | None = None,
) -> dict[str, Path]:
    """Copy both templates and replace only evidence-backed audit inputs."""
    output_dir.mkdir(parents=True, exist_ok=True)
    source_rows = _source_rows(document_name, pages)
    project_metadata = dict(project_metadata or {})
    preferred_scope = str(project_metadata.get("financial_scope") or "母公司")
    basis_date = _basis_date(source_rows)
    if basis_date:
        project_metadata["report_date"] = basis_date
    supporting_materials = supporting_materials or []
    outputs = {
        "reporting_workbook": output_dir / "资产基础法_资产清查.xlsx",
        "income_workbook": output_dir / "收益法_市场法.xlsx",
        "trace_workbook": output_dir / "审计材料_规则与溯源.xlsx",
    }
    shutil.copy2(asset_template, outputs["reporting_workbook"])
    shutil.copy2(income_template, outputs["income_workbook"])

    asset = load_workbook(outputs["reporting_workbook"])
    income = load_workbook(outputs["income_workbook"])
    if progress_callback:
        progress_callback("load_mapping_rules", "completed", "模板结构、标准科目别名和公式坐标规则已加载", 64)
    income_value_bounds = _nonempty_bounds(income)
    _apply_project_identity(asset, income, project_metadata, basis_date)
    asset_rows: list[list[Any]] = []
    _clear_asset_sample_data(asset)
    if progress_callback:
        progress_callback("map_asset", "running", "正在按资产法模板逐单元格匹配审定账面值", 64)
    _fill_asset_template(asset, source_rows, asset_rows, preferred_scope=preferred_scope)
    if progress_callback:
        progress_callback("map_asset", "completed", f"资产法映射完成，共记录 {len(asset_rows)} 条单元格规则", 72)

    income_rows: list[list[Any]] = []
    history = {"历资表", "历利表", "历现表"}
    # The B:E historical columns in the reference file contain Tongfu's
    # sample figures.  They must be cleared first; only D/E are re-populated
    # from audited current/prior values below.  Keeping those numbers would
    # silently mix companies in downstream forecast formulas.
    _clear_nonformula_numbers(income, {"目录1", "目录2", "项目信息"})
    _clear_income_sample_text(income)
    project = income["项目信息"]
    project["B8"] = "万元"
    if progress_callback:
        progress_callback("map_income", "running", "正在按收益法模板填入审定历史三表并保留公式链", 73)
    history_categories = {"历资表": "资产负债表", "历利表": "利润表", "历现表": "现金流量表"}
    history_periods: dict[str, list[str]] = {}
    for name in history:
        history_periods[name] = _fill_history(
            income,
            name,
            history_categories[name],
            source_rows,
            income_rows,
            preferred_scope=preferred_scope,
        )
    model_periods = history_periods.get("历利表") or history_periods.get("历资表") or []
    _apply_model_period_headers(income, model_periods, basis_date)
    enhance_from_notes(
        asset=asset,
        income=income,
        document_name=document_name,
        pages=pages,
        asset_mappings=asset_rows,
        income_mappings=income_rows,
        scope=preferred_scope,
        statement_source_rows=source_rows,
    )
    if progress_callback:
        progress_callback("map_income", "completed", f"收益法映射完成，共记录 {len(income_rows)} 条单元格规则", 82)
    _set_recalculation(asset)
    _set_recalculation(income)
    asset.save(outputs["reporting_workbook"])
    income.save(outputs["income_workbook"])

    # Keep the two client-facing workbooks structurally identical to the
    # supplied Tongfu templates.  The traceability ledger is deliberately a
    # separate audit working paper, not extra tabs inserted into a template.
    trace = Workbook()
    default_sheet = trace.active
    trace.remove(default_sheet)
    _add_combined_mapping_sheet(trace, asset_rows, income_rows)
    _add_formula_rules_from(trace, [
        ("资产基础法_资产清查.xlsx", asset),
        ("收益法_市场法.xlsx", income),
    ])
    _add_audit_source_sheets(trace, document_name, pages)
    _add_supporting_material_sheets(trace, supporting_materials)
    _write_metadata_sheet(trace, project_metadata, document_name)
    _add_template_field_registry(trace, asset, income, asset_rows, income_rows, income_value_bounds)
    _add_note_candidates(trace, document_name, pages)
    _set_recalculation(trace)
    trace.save(outputs["trace_workbook"])
    if progress_callback:
        progress_callback("build_rule_graph", "running", "正在生成每个 Sheet、单元格、来源和公式依赖的通用规则图", 88)
    _write_rule_graph(output_dir, document_name, asset_rows, income_rows, supporting_materials)
    build_rule_graph_bundle(
        output_dir=output_dir / "通用逐单元格规则图",
        workbook_paths={
            "资产基础法_资产清查.xlsx": outputs["reporting_workbook"],
            "收益法_市场法.xlsx": outputs["income_workbook"],
        },
        trace_path=outputs["trace_workbook"],
        case_name=str(project_metadata.get("company_name") or document_name),
    )
    if progress_callback:
        progress_callback("build_rule_graph", "completed", "逐单元格规则图和待补资料清单已生成", 94)
    return outputs
