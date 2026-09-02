"""Turn heterogeneous audit PDFs into auditable role workbooks.

The previous workflow expected an asset/reporting workbook and an income/market
workbook to be supplied by a human.  This module deliberately does *not*
pretend that an audit report contains an appraisal conclusion.  It extracts
only tables that the report actually proves, preserves every source cell, and
marks valuation-only cells as unresolved for the later appraisal workpapers.
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from datetime import date
from pathlib import Path
import json
import re
from typing import Any, Iterable

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


ROLE_REPORTING = "reporting_workbook"
ROLE_INCOME = "income_workbook"
ROLE_LABELS = {
    ROLE_REPORTING: "资产基础法_资产清查",
    ROLE_INCOME: "收益法_市场法",
}

TABLE_RULES: tuple[tuple[str, tuple[str, ...], tuple[str, ...]], ...] = (
    ("资产负债表", ("资产负债表",), (ROLE_REPORTING,)),
    ("利润表", ("利润表", "损益表"), (ROLE_INCOME,)),
    ("现金流量表", ("现金流量表",), (ROLE_INCOME,)),
    ("所有者权益变动表", ("所有者权益变动表", "股东权益变动表"), (ROLE_REPORTING,)),
    (
        "资产明细",
        ("固定资产", "无形资产", "长期股权投资", "存货", "资产清查"),
        (ROLE_REPORTING,),
    ),
    (
        "资产评估结果",
        ("资产基础法", "资产评估结果", "评估价值", "评估汇总"),
        (ROLE_REPORTING,),
    ),
    (
        "收益法估值",
        ("收益法", "自由现金流", "折现率", "股东全部权益价值"),
        (ROLE_INCOME,),
    ),
    (
        "市场法估值",
        ("市场法", "可比公司", "市盈率", "市净率"),
        (ROLE_INCOME,),
    ),
)


@dataclass(frozen=True)
class ExtractedTable:
    source_file: str
    page_number: int
    table_id: str
    category: str
    roles: tuple[str, ...]
    matrix: list[list[str]]
    period_headers: dict[int, str]
    confidence: str


def _compact(value: Any) -> str:
    return re.sub(r"[\s:：()（）]", "", str(value or "")).lower()


def _cell_value(value: Any) -> Any:
    """Keep text unless it is unambiguously a standalone number."""
    text = str(value or "").strip()
    if not text:
        return ""
    normalized = text.translate(str.maketrans("，．－＋１２３４５６７８９０", ",.-+1234567890"))
    candidate = normalized.replace(",", "")
    if re.fullmatch(r"[+-]?\d+(?:\.\d+)?", candidate):
        try:
            return float(candidate) if "." in candidate else int(candidate)
        except ValueError:
            pass
    return text


def _page_text(page: dict[str, Any]) -> str:
    blocks = page.get("blocks", []) if isinstance(page, dict) else []
    block_text = " ".join(str(item.get("text", "")) for item in blocks if isinstance(item, dict))
    table_text = " ".join(
        str(cell.get("text", ""))
        for table in page.get("tables", []) if isinstance(table, dict)
        for cell in table.get("cells", []) if isinstance(cell, dict)
    )
    return f"{block_text} {table_text}"


def _table_category(page: dict[str, Any], cells: list[dict[str, Any]]) -> tuple[str, tuple[str, ...]]:
    # A page may mention a financial statement while its only table is an
    # audit-procedure narrative.  Classify from the *table itself*, with a
    # combination of accounting labels rather than a single page keyword.
    text = _compact(" ".join(str(cell.get("text", "")) for cell in cells))
    numeric_count = sum(isinstance(_cell_value(cell.get("text")), (int, float)) for cell in cells)
    if len(cells) < 5 or numeric_count < 2:
        return "未分类财务表", ()
    if "资产负债表" in text or (
        any(token in text for token in ("资产总计", "资产合计", "负债合计", "负债总计"))
        and any(token in text for token in ("项目", "期末余额", "流动资产", "流动负债", "所有者权益"))
    ):
        return "资产负债表", (ROLE_REPORTING,)
    if "利润表" in text or sum(token in text for token in ("营业收入", "营业成本", "净利润", "利润总额")) >= 2:
        return "利润表", (ROLE_INCOME,)
    if "现金流量表" in text or sum(token in text for token in ("经营活动", "投资活动", "筹资活动", "现金及现金等价物")) >= 2:
        return "现金流量表", (ROLE_INCOME,)
    if "所有者权益变动表" in text or (
        "实收资本" in text and ("未分配利润" in text or "盈余公积" in text)
    ):
        return "所有者权益变动表", (ROLE_REPORTING,)
    if "资产基础法" in text or ("账面价值" in text and "评估价值" in text and "增值率" in text):
        return "资产评估结果", (ROLE_REPORTING,)
    if "收益法" in text or ("自由现金流" in text and "折现率" in text):
        return "收益法估值", (ROLE_INCOME,)
    if "市场法" in text or ("可比公司" in text and ("市盈率" in text or "市净率" in text)):
        return "市场法估值", (ROLE_INCOME,)
    if sum(token in text for token in ("固定资产", "无形资产", "长期股权投资", "存货")) >= 2 or (
        "固定资产" in text and ("原值" in text or "累计折旧" in text or "账面价值" in text)
    ):
        return "资产明细", (ROLE_REPORTING,)
    return "未分类财务表", ()


def _matrix(cells: Iterable[dict[str, Any]]) -> list[list[str]]:
    cells = list(cells)
    max_row = max((int(cell.get("row") or 0) for cell in cells), default=0)
    max_col = max((int(cell.get("column") or 0) for cell in cells), default=0)
    matrix = [["" for _ in range(max_col)] for _ in range(max_row)]
    for cell in cells:
        row, column = int(cell.get("row") or 0), int(cell.get("column") or 0)
        if row > 0 and column > 0:
            matrix[row - 1][column - 1] = str(cell.get("text") or "")
    return matrix


def _reporting_date(page: dict[str, Any]) -> tuple[int, int, int] | None:
    text = " ".join(str(block.get("text", "")) for block in page.get("blocks", []) if isinstance(block, dict))
    text = re.sub(r"(年\s*十二月\s*)三十白", r"\1三十一日", text)
    match = re.search(r"((?:19|20)\d{2})\s*年\s*(\d{1,2})\s*[月角]\s*(\d{1,2})\s*日?", text)
    if match:
        year, month, day = (int(value) for value in match.groups())
        try:
            date(year, month, day)
            return year, month, day
        except ValueError:
            pass

    digits = {"〇": 0, "零": 0, "一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}

    def chinese_number(value: str) -> int | None:
        value = value.strip()
        if value == "十":
            return 10
        if "十" in value:
            left, right = value.split("十", 1)
            tens = digits.get(left, 1) if left else 1
            ones = digits.get(right, 0) if right else 0
            return tens * 10 + ones
        if value and all(character in digits for character in value):
            return int("".join(str(digits[character]) for character in value))
        return None

    chinese = re.search(r"([〇零一二三四五六七八九]{4})\s*年\s*([一二三四五六七八九十]{1,3})\s*月\s*([一二三四五六七八九十]{1,3})\s*日?", text)
    if chinese:
        year = chinese_number(chinese.group(1)); month = chinese_number(chinese.group(2)); day = chinese_number(chinese.group(3))
        if year and month and day:
            try:
                date(year, month, day)
                return year, month, day
            except ValueError:
                pass

    annual = re.search(r"((?:19|20)\d{2})\s*年度", text)
    if annual:
        return int(annual.group(1)), 12, 31
    chinese_annual = re.search(r"([〇零一二三四五六七八九]{4})\s*年度", text)
    year = chinese_number(chinese_annual.group(1)) if chinese_annual else None
    return (year, 12, 31) if year else None


def _period_headers(matrix: list[list[str]], report_date: tuple[int, int, int] | None, page_text: str = "") -> dict[int, str]:
    if report_date is None:
        return {}
    year, month, day = report_date
    result: dict[int, str] = {}
    for row in matrix[:4]:
        for index, value in enumerate(row, 1):
            label = _compact(value)
            if any(token in label for token in ("上年年末", "上期期末", "年初余额", "期初余额")):
                result[index] = f"{year - 1:04d}-{month:02d}-{day:02d}"
            elif any(token in label for token in ("期末余额", "本期期末", "本年年末", "年末余额")):
                result[index] = f"{year:04d}-{month:02d}-{day:02d}"
            elif any(token in label for token in ("上期金额", "上期数", "上年同期", "上年金额", "上年累计")):
                result[index] = f"{year - 1:04d}年度"
            elif any(token in label for token in ("本期金额", "本期数", "本年累计", "本年金额")):
                result[index] = f"{year:04d}年度"
    if result:
        return result

    # Some scan OCR keeps the statement headings as page text while the table
    # starts directly at the first account row.  Infer columns only when the
    # page explicitly names current/prior periods and exactly two columns have
    # repeated money-like values; note/line-number columns are thereby excluded.
    compact_page = _compact(page_text)
    current_tokens = ("年末余额", "期末余额", "本年金额", "本期金额", "本年累计")
    prior_tokens = ("年初余额", "期初余额", "上年金额", "上期金额", "上年累计", "上年年末")
    current_position = min((compact_page.find(token) for token in current_tokens if token in compact_page), default=-1)
    prior_position = min((compact_page.find(token) for token in prior_tokens if token in compact_page), default=-1)
    if current_position < 0 or prior_position < 0:
        return result
    money_scores: dict[int, int] = {}
    for row in matrix:
        for column, raw in enumerate(row, 1):
            parsed = _cell_value(raw)
            text = str(raw or "").translate(str.maketrans("，．－＋", ",.-+"))
            if isinstance(parsed, (int, float)) and (abs(float(parsed)) >= 100 or bool(re.search(r"[,.]\d{1,2}$", text))):
                money_scores[column] = money_scores.get(column, 0) + 1
    threshold = max(3, len(matrix) // 8)
    amount_columns = sorted(column for column, score in money_scores.items() if score >= threshold)
    if len(amount_columns) != 2:
        return result
    current_period = f"{year:04d}-{month:02d}-{day:02d}"
    prior_period = f"{year - 1:04d}-{month:02d}-{day:02d}"
    ordered_periods = (current_period, prior_period) if current_position < prior_position else (prior_period, current_period)
    return dict(zip(amount_columns, ordered_periods))


def tables_from_pages(source_file: str, pages: list[dict[str, Any]]) -> list[ExtractedTable]:
    """Classify OCR/PDF tables solely from their contents, not their filenames."""
    result: list[ExtractedTable] = []
    page_dates = {
        int(page.get("page_number") or 0): _reporting_date(page)
        for page in pages
        if any(token in _page_text(page) for token in ("资产负债表", "利润表", "现金流量表"))
    }
    for page in pages:
        page_number = int(page.get("page_number") or 0)
        report_date = _reporting_date(page)
        if report_date is None:
            nearby = [
                (abs(other_page - page_number), other_date)
                for other_page, other_date in page_dates.items()
                if other_date is not None and abs(other_page - page_number) <= 3
            ]
            if nearby:
                report_date = min(nearby, key=lambda item: item[0])[1]
        for index, table in enumerate(page.get("tables", []) or [], 1):
            cells = [cell for cell in table.get("cells", []) if isinstance(cell, dict)]
            if not cells:
                continue
            category, roles = _table_category(page, cells)
            result.append(
                ExtractedTable(
                    source_file=source_file,
                    page_number=page_number,
                    table_id=str(table.get("table_id") or f"p{page_number}-t{index}"),
                    category=category,
                    roles=roles,
                    matrix=_matrix(cells),
                    period_headers=_period_headers(_matrix(cells), report_date, _page_text(page)),
                    confidence="材料直接证明" if roles else "待人工确认",
                )
            )
    return result


def needs_ocr(pages: list[dict[str, Any]]) -> bool:
    """OCR is needed only when local PDF extraction found no usable text/table."""
    return not any(_page_text(page).strip() or page.get("tables") for page in pages)


def _safe_sheet_title(title: str, used: set[str]) -> str:
    base = re.sub(r"[\\/*?:\[\]]", "_", title)[:31] or "财务表"
    candidate, number = base, 2
    while candidate in used:
        suffix = f"_{number}"
        candidate = f"{base[:31-len(suffix)]}{suffix}"
        number += 1
    used.add(candidate)
    return candidate


def _style(sheet) -> None:
    fill = PatternFill("solid", fgColor="1F4E78")
    for cell in sheet[1]:
        cell.fill = fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    sheet.freeze_panes = "A2"
    sheet.row_dimensions[1].height = 26
    for column in range(1, max(sheet.max_column, 1) + 1):
        values = [str(sheet.cell(row, column).value or "") for row in range(1, sheet.max_row + 1)]
        sheet.column_dimensions[get_column_letter(column)].width = min(max(max(map(len, values), default=8) + 2, 10), 36)
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)


def _add_metadata(workbook: Workbook, document_name: str, role: str, tables: list[ExtractedTable]) -> None:
    sheet = workbook.create_sheet("生成说明")
    sheet.append(["项目", "内容"])
    sheet.append(["工作簿角色", ROLE_LABELS[role]])
    sheet.append(["来源审计材料", document_name])
    sheet.append(["生成原则", "仅写入材料直接证明的数值；评估值、折现率等不能从审计报表推导的内容保持待补充。"])
    sheet.append(["已归类表数", len(tables)])
    _style(sheet)


def _add_mapping_sheet(workbook: Workbook, rows: list[list[Any]]) -> None:
    sheet = workbook.create_sheet("映射规则")
    sheet.append(["目标工作表", "目标单元格", "源文件", "源页码", "源表格", "源行", "源列", "映射方式", "计算表达式", "证据等级"])
    for row in rows:
        sheet.append(row)
    _style(sheet)
    sheet.auto_filter.ref = f"A1:J{max(sheet.max_row, 1)}"


def _write_table(workbook: Workbook, table: ExtractedTable, used: set[str], mapping_rows: list[list[Any]]) -> str:
    title = _safe_sheet_title(f"{table.category}_{table.page_number}", used)
    sheet = workbook.create_sheet(title)
    sheet.append([table.category, f"来源：{table.source_file} 第{table.page_number}页，表格 {table.table_id}"])
    for column, period in table.period_headers.items():
        sheet.cell(row=2, column=column).value = period
        mapping_rows.append(
            [
                title,
                sheet.cell(row=2, column=column).coordinate,
                table.source_file,
                table.page_number,
                table.table_id,
                "页头",
                column,
                "页头期间归一化",
                f"PERIOD({period})",
                "材料直接证明",
            ]
        )
    for row_index, row in enumerate(table.matrix, 1):
        for column_index, value in enumerate(row, 1):
            target = sheet.cell(row=row_index + 2, column=column_index)
            target.value = _cell_value(value)
            if str(value).strip():
                mapping_rows.append(
                    [
                        title,
                        target.coordinate,
                        table.source_file,
                        table.page_number,
                        table.table_id,
                        row_index,
                        column_index,
                        "源值直取",
                        f"DIRECT(源表[{row_index},{column_index}])",
                        table.confidence,
                    ]
                )
    _style(sheet)
    return title


def _row_numeric_cells(sheet, aliases: tuple[str, ...]) -> list[str]:
    """Locate a statement line and return its numeric cells in display order."""
    compact_aliases = {_compact(alias) for alias in aliases}
    for row in sheet.iter_rows(min_row=3):
        labels = {_compact(cell.value) for cell in row if isinstance(cell.value, str)}
        if not labels.intersection(compact_aliases):
            continue
        values = [cell.coordinate for cell in row if isinstance(cell.value, (int, float))]
        if values:
            return values
    return []


def _add_accounting_checks(workbook: Workbook, mapping_rows: list[list[Any]]) -> None:
    """Add only accounting identities that can be verified from extracted cells.

    These formulas are not estimates and do not create appraisal values.  They
    make the inferred workbook logic explicit and flag OCR/layout errors.
    """
    checks: list[tuple[str, str, str, str]] = []
    definitions = (
        ("资产总计 = 流动资产合计 + 非流动资产合计", ("流动资产合计",), ("非流动资产合计",), ("资产总计",)),
        ("负债合计 = 流动负债合计 + 非流动负债合计", ("流动负债合计",), ("非流动负债合计",), ("负债合计", "负债总计")),
    )
    raw_prefixes = ("资产负债表_", "利润表_", "现金流量表_")
    source_sheets = [sheet for sheet in workbook.worksheets if sheet.title.startswith(raw_prefixes)]
    for sheet in source_sheets:
        if sheet.title in {"生成说明", "映射规则", "待补充"}:
            continue
        for label, left_aliases, right_aliases, total_aliases in definitions:
            left = _row_numeric_cells(sheet, left_aliases)
            right = _row_numeric_cells(sheet, right_aliases)
            total = _row_numeric_cells(sheet, total_aliases)
            for first, second, expected in zip(left, right, total, strict=False):
                checks.append((label, sheet.title, first, second + "|" + expected))
    if not checks:
        return
    sheet = workbook.create_sheet("核验计算")
    sheet.append(["校验项", "工作表", "公式", "差额", "状态", "规则等级"])
    for row_number, (label, source_sheet, first, rest) in enumerate(checks, 2):
        second, expected = rest.split("|", 1)
        formula = f"='{source_sheet}'!{first}+'{source_sheet}'!{second}-'{source_sheet}'!{expected}"
        sheet.cell(row_number, 1).value = label
        sheet.cell(row_number, 2).value = source_sheet
        sheet.cell(row_number, 3).value = formula
        sheet.cell(row_number, 4).value = f"=C{row_number}"
        sheet.cell(row_number, 5).value = f'=IF(ABS(D{row_number})<0.01,"通过","需复核")'
        sheet.cell(row_number, 6).value = "会计恒等式可验证"
        mapping_rows.append(
            [
                sheet.title,
                f"C{row_number}:E{row_number}",
                "生成工作簿",
                "",
                source_sheet,
                "",
                ",".join((first, second, expected)),
                "会计恒等式校验",
                formula,
                "会计恒等式可验证",
            ]
        )
    _style(sheet)


# These are the business workpaper rows the existing report mapping expects.
# They intentionally describe a *generic* appraisal workbook, rather than
# copying one client's private layout.  Rows supported by the audit report are
# populated; valuation-only inputs remain visible and explicitly unresolved.
BALANCE_ROWS = (
    "货币资金", "交易性金融资产", "应收票据", "应收账款", "预付款项", "其他应收款", "存货",
    "合同资产", "其他流动资产", "流动资产合计", "长期股权投资", "投资性房地产", "固定资产",
    "在建工程", "使用权资产", "无形资产", "长期待摊费用", "递延所得税资产", "其他非流动资产",
    "非流动资产合计", "资产总计", "短期借款", "应付票据", "应付账款", "合同负债", "应付职工薪酬",
    "应交税费", "其他应付款", "一年内到期的非流动负债", "其他流动负债", "流动负债合计",
    "长期借款", "租赁负债", "长期应付款", "预计负债", "递延所得税负债", "其他非流动负债",
    "非流动负债合计", "负债合计", "实收资本", "资本公积", "盈余公积", "未分配利润", "所有者权益合计",
)

INCOME_ROWS = (
    "一、营业收入", "减：营业成本", "税金及附加", "销售费用", "管理费用", "研发费用", "财务费用",
    "加：其他收益", "投资收益", "公允价值变动收益", "信用减值损失", "资产减值损失", "资产处置收益",
    "二、营业利润", "加：营业外收入", "减：营业外支出", "三、利润总额", "减：所得税费用", "四、净利润",
)

CASHFLOW_ROWS = (
    "经营活动产生的现金流量净额", "投资活动产生的现金流量净额", "筹资活动产生的现金流量净额",
    "汇率变动对现金及现金等价物的影响", "现金及现金等价物净增加额", "期初现金及现金等价物余额", "期末现金及现金等价物余额",
)

DETAIL_SCHEDULES = (
    ("表3-1_货币资金", "货币资金", ("银行", "账户", "币种", "账面余额", "清查金额", "评估价值", "备注")),
    ("表3-2_应收票据", "应收票据", ("出票人/承兑人", "票据编号", "到期日", "账面余额", "坏账准备", "评估价值", "备注")),
    ("表3-3_应收账款", "应收账款", ("客户名称", "账龄", "账面余额", "坏账准备", "账面净额", "评估价值", "备注")),
    ("表3-4_预付款项", "预付款项", ("单位名称", "款项性质", "账龄", "账面余额", "评估价值", "备注")),
    ("表3-5_其他应收款", "其他应收款", ("单位/个人", "款项性质", "账龄", "账面余额", "坏账准备", "评估价值", "备注")),
    ("表3-9_存货", "存货", ("存货类别", "名称", "数量", "单位", "账面余额", "跌价准备", "评估价值", "备注")),
    ("表4-1_长期股权投资", "长期股权投资", ("被投资单位", "持股比例", "账面价值", "评估价值", "评估方法", "备注")),
    ("表4-2_投资性房地产", "投资性房地产", ("资产名称", "坐落/位置", "面积", "账面价值", "评估价值", "备注")),
    ("表4-6_固定资产", "固定资产", ("资产编号", "资产名称", "类别", "原值", "累计折旧", "账面净值", "评估价值", "备注")),
    ("表4-12_无形资产", "无形资产", ("资产名称", "权利类型", "取得日期", "账面价值", "评估价值", "评估方法", "备注")),
    ("表4-13_长期待摊费用", "长期待摊费用", ("项目", "受益期限", "账面价值", "评估价值", "备注")),
)


def _canonical_label(value: Any) -> str:
    label = re.sub(r"^\d+[、.．]", "", str(value or "").strip())
    label = re.sub(r"[（(]注?\d+[）)]", "", label)
    label = re.sub(r"\s+", "", label)
    aliases = {
        "资产合计": "资产总计", "负债总计": "负债合计", "所有者权益合计": "所有者权益合计",
        "股东权益合计": "所有者权益合计", "营业收入": "一、营业收入", "营业成本": "减：营业成本",
        "营业利润": "二、营业利润", "利润总额": "三、利润总额", "净利润": "四、净利润",
        "所得税费用": "减：所得税费用", "其他收益": "加：其他收益", "投资收益": "投资收益",
        "营业外收入": "加：营业外收入", "营业外支出": "减：营业外支出",
    }
    label = aliases.get(label, label)
    compact = _compact(label)
    for target in (*BALANCE_ROWS, *INCOME_ROWS, *CASHFLOW_ROWS):
        target_compact = _compact(target)
        if compact == target_compact or compact.startswith(target_compact):
            return target
    # Common statutory-statement variants which carry parenthetical wording.
    if "所有者权益" in label or "股东权益" in label:
        return "所有者权益合计" if "合计" in label or "总计" in label else label
    if label.startswith("实收资本") or label.startswith("股本"):
        return "实收资本"
    return label


def _ocr_number(value: Any) -> Any:
    """Normalize only unmistakable OCR number punctuation, never estimate a value."""
    if isinstance(value, (int, float)):
        return value
    raw_text = str(value or "").strip().replace("，", ",")
    # When OCR renders the decimal point as a comma, e.g.
    # ``16，281，363，74``, `_cell_value` would otherwise delete every comma
    # and inflate the amount by 100.  A final two-digit group is an
    # unambiguous cents field when the token has at least two separators.
    raw_marks = [index for index, char in enumerate(raw_text) if char in ".,"]
    if len(raw_marks) >= 2 and re.fullmatch(r"[+-]?[\d.,]+", raw_text):
        last = raw_marks[-1]
        decimal = raw_text[last + 1 :]
        if decimal and len(decimal) <= 2:
            rebuilt = re.sub(r"[.,]", "", raw_text[:last]) + "." + decimal
            try:
                return float(rebuilt)
            except ValueError:
                pass
    direct = _cell_value(value)
    if isinstance(direct, (int, float)):
        return direct
    text = raw_text
    # OCR engines occasionally substitute visually similar Latin glyphs in
    # an otherwise unambiguous numeric token (for example ``2,29S,884.83``).
    # Apply the correction only when the complete corrected token is numeric;
    # ordinary words therefore remain untouched and no amount is estimated.
    corrected = text.translate(str.maketrans({"G": "6", "g": "6", "O": "0", "o": "0", "I": "1", "l": "1"}))
    if re.fullmatch(r"[+-]?[\d.,]+", corrected) and re.search(r"\d", text):
        text = corrected
    if not re.fullmatch(r"[+-]?[\d.,]+", text):
        return direct
    marks = [index for index, char in enumerate(text) if char in ".,"]
    if len(marks) < 2:
        return direct
    last = marks[-1]
    decimal = text[last + 1 :]
    if not decimal or len(decimal) > 2:
        return direct
    rebuilt = re.sub(r"[.,]", "", text[:last]) + "." + decimal
    try:
        return float(rebuilt)
    except ValueError:
        return direct


def _statement_sources(tables: list[ExtractedTable], category: str) -> dict[str, tuple[ExtractedTable, int, int, Any]]:
    """Return one evidence-backed value per canonical statement line.

    A statement frequently continues over several OCR tables/pages.  Start
    with the most complete table and add only *missing* labels from the
    remaining tables.  If a group/parent report repeats a label, the selected
    primary table remains authoritative and every retained value still has its
    own page/table evidence in the mapping sheet.
    """
    candidates: list[tuple[int, ExtractedTable, dict[str, tuple[int, int, Any]]]] = []
    for table in tables:
        if table.category != category:
            continue
        current: dict[str, tuple[int, int, Any]] = {}
        for row_number, row in enumerate(table.matrix, 1):
            label = next((cell for cell in row if str(cell or "").strip()), "")
            key = _canonical_label(label)
            if not key:
                continue
            for column_number, raw in enumerate(row, 1):
                parsed = _ocr_number(raw)
                if isinstance(parsed, (int, float)):
                    current.setdefault(key, (row_number, column_number, parsed))
                    break
        candidates.append((len(current), table, current))
    if not candidates:
        return {}
    candidates.sort(key=lambda item: item[0], reverse=True)
    result: dict[str, tuple[ExtractedTable, int, int, Any]] = {}
    for _, table, values in candidates:
        for key, (row, column, value) in values.items():
            result.setdefault(key, (table, row, column, value))
    return result


def _input_fill(cell) -> None:
    cell.fill = PatternFill("solid", fgColor="FFF2CC")
    cell.font = Font(color="0000FF")


def _section(sheet, row: int, title: str, end_column: int) -> None:
    sheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=end_column)
    cell = sheet.cell(row, 1, title)
    cell.fill = PatternFill("solid", fgColor="1F4E78")
    cell.font = Font(color="FFFFFF", bold=True)
    cell.alignment = Alignment(horizontal="left")


def _map_direct(mapping_rows: list[list[Any]], sheet_name: str, cell, source: tuple[ExtractedTable, int, int, Any], note: str = "") -> None:
    table, row, column, _ = source
    mapping_rows.append([
        sheet_name, cell.coordinate, table.source_file, table.page_number, table.table_id,
        row, column, "源值直取" if not note else note,
        f"DIRECT(源表[{row},{column}])", table.confidence,
    ])


def _map_pending(mapping_rows: list[list[Any]], sheet_name: str, cell, reason: str) -> None:
    mapping_rows.append([sheet_name, cell.coordinate, "", "", "", "", "", "待评估师补充", reason, "材料未提供"])


def _add_asset_workpapers(workbook: Workbook, tables: list[ExtractedTable], mappings: list[list[Any]]) -> None:
    sources = _statement_sources(tables, "资产负债表")
    overview = workbook.create_sheet("项目说明", 1)
    overview.append(["资产基础法 / 资产清查工作簿"])
    overview.append(["用途", "完整资产评估底稿结构；审计材料可证明的账面数据已导入，评估参数和评估值须由评估师补充。"])
    overview.append(["颜色", "黄色蓝字：待填输入；绿色：跨表取数；黑字：计算公式；映射规则表记录每个来源或公式。"])
    overview.append(["来源", "审计报告及其附注；不以审计报告反推评估价值。"])
    _style(overview)

    summary = workbook.create_sheet("表1_资产评估结果汇总表", 2)
    summary.append(["资产评估结果汇总表（单位：按原报告单位；评估价值待补充）"])
    summary.append(["项目", "账面价值", "评估价值", "增减值", "增值率", "状态", "账面来源"])
    summary_rows = [
        "流动资产合计", "非流动资产合计", "资产总计", "流动负债合计", "非流动负债合计", "负债合计", "所有者权益合计",
    ]
    for index, label in enumerate(summary_rows, 3):
        summary.cell(index, 1, label)
        source = sources.get(label)
        if source:
            summary.cell(index, 2, source[3])
            _map_direct(mappings, summary.title, summary.cell(index, 2), source)
            summary.cell(index, 7, f"{source[0].source_file} 第{source[0].page_number}页")
        else:
            _input_fill(summary.cell(index, 2))
            _map_pending(mappings, summary.title, summary.cell(index, 2), "审计材料未识别到该账面科目")
        _input_fill(summary.cell(index, 3))
        _map_pending(mappings, summary.title, summary.cell(index, 3), "资产清查/评估作业结果")
        summary.cell(index, 4, f'=IF(OR(B{index}="",C{index}=""),"",C{index}-B{index})')
        summary.cell(index, 5, f'=IFERROR(IF(D{index}="","",D{index}/B{index}),"")')
        summary.cell(index, 6, f'=IF(C{index}="","待评估","已填")')
        mappings.append([summary.title, f"D{index}:F{index}", "", "", "", "", "", "公式计算", f"增减值=C{index}-B{index}; 增值率=D{index}/B{index}", "确定性公式"])
    summary.cell(10, 1, "净资产（所有者权益）")
    summary.cell(10, 2, "=B5-B8")
    summary.cell(10, 3, "=C5-C8")
    summary.cell(10, 4, "=C10-B10")
    summary.cell(10, 5, '=IFERROR(D10/B10,"")')
    summary.cell(10, 6, '=IF(C10="","待评估","已填")')
    mappings.append([summary.title, "B10:F10", "表1_资产评估结果汇总表", "", "", "", "B5,B8,C5,C8", "公式计算", "净资产=资产总计-负债合计", "确定性公式"])
    _style(summary)

    scope = workbook.create_sheet("表2_资产负债范围表", 3)
    scope.append(["资产负债范围表（审计账面数；单位按原报告）"])
    scope.append(["科目", "账面金额", "是否纳入评估范围", "清查/权属资料", "评估价值", "差异说明", "账面来源"])
    for row_number, label in enumerate(BALANCE_ROWS, 3):
        scope.cell(row_number, 1, label)
        source = sources.get(label)
        if source:
            scope.cell(row_number, 2, source[3])
            scope.cell(row_number, 7, f"{source[0].source_file} 第{source[0].page_number}页")
            _map_direct(mappings, scope.title, scope.cell(row_number, 2), source)
        else:
            _input_fill(scope.cell(row_number, 2))
            _map_pending(mappings, scope.title, scope.cell(row_number, 2), "审计材料未识别或该科目不适用")
        scope.cell(row_number, 3, "待判断")
        scope.cell(row_number, 4, "待补充")
        _input_fill(scope.cell(row_number, 3)); _input_fill(scope.cell(row_number, 4)); _input_fill(scope.cell(row_number, 5))
        _map_pending(mappings, scope.title, scope.cell(row_number, 3), "评估范围判断")
        _map_pending(mappings, scope.title, scope.cell(row_number, 5), "资产评估结果")
    _style(scope)

    for sheet_name, account, headers in DETAIL_SCHEDULES:
        sheet = workbook.create_sheet(sheet_name)
        sheet.append([f"{account}清查评估明细表"])
        sheet.append(["审计报表对应账面余额", ""])
        source = sources.get(account)
        if source:
            sheet.cell(2, 2, source[3])
            _map_direct(mappings, sheet.title, sheet.cell(2, 2), source)
        else:
            _input_fill(sheet.cell(2, 2))
            _map_pending(mappings, sheet.title, sheet.cell(2, 2), "审计报表未提供该明细科目")
        sheet.append(list(headers))
        for column in range(1, len(headers) + 1):
            _input_fill(sheet.cell(4, column))
            _map_pending(mappings, sheet.title, sheet.cell(4, column), "资产清查明细/评估师工作底稿")
        _style(sheet)


def _add_income_workpapers(workbook: Workbook, tables: list[ExtractedTable], mappings: list[list[Any]]) -> None:
    balance_sources = _statement_sources(tables, "资产负债表")
    income_sources = _statement_sources(tables, "利润表")
    cash_sources = _statement_sources(tables, "现金流量表")
    project = workbook.create_sheet("项目信息", 1)
    project.append(["收益法 / 市场法工作簿"])
    for label in ("被评估企业", "评估基准日", "价值类型", "币种/单位", "模型版本", "来源说明"):
        project.append([label, "待补充"])
        _input_fill(project.cell(project.max_row, 2))
        _map_pending(mappings, project.title, project.cell(project.max_row, 2), "项目设定或非财务基础资料")
    _style(project)

    def add_history(sheet_name: str, title: str, labels: tuple[str, ...], sources: dict[str, tuple[ExtractedTable, int, int, Any]]) -> None:
        sheet = workbook.create_sheet(sheet_name)
        sheet.append([title])
        sheet.append(["项目", "本期实际", "上期实际", "来源/说明"])
        for row_number, label in enumerate(labels, 3):
            sheet.cell(row_number, 1, label)
            source = sources.get(label)
            if source:
                sheet.cell(row_number, 2, source[3])
                sheet.cell(row_number, 4, f"{source[0].source_file} 第{source[0].page_number}页")
                _map_direct(mappings, sheet.title, sheet.cell(row_number, 2), source)
            else:
                _input_fill(sheet.cell(row_number, 2))
                _map_pending(mappings, sheet.title, sheet.cell(row_number, 2), "审计材料未识别到该历史科目")
            _input_fill(sheet.cell(row_number, 3))
            _map_pending(mappings, sheet.title, sheet.cell(row_number, 3), "需从同口径上期报表取数")
        _style(sheet)

    add_history("历资表", "历史资产负债表（单位按原报告）", BALANCE_ROWS, balance_sources)
    add_history("历利表", "历史利润表（单位按原报告）", INCOME_ROWS, income_sources)
    add_history("历现表", "历史现金流量表（单位按原报告）", CASHFLOW_ROWS, cash_sources)

    forecast = workbook.create_sheet("收益预测")
    forecast.append(["收益预测（黄色蓝字为需评估师确认的假设；单位按项目设定）"])
    forecast.append(["项目", "历史最近期", "预测第1年", "预测第2年", "预测第3年", "预测第4年", "预测第5年", "假设/来源"])
    for row_number, label in enumerate(("营业收入", "收入增长率", "营业成本", "毛利率", "销售费用", "管理费用", "研发费用", "财务费用", "所得税费用", "净利润"), 3):
        forecast.cell(row_number, 1, label)
        if label == "营业收入" and income_sources.get("一、营业收入"):
            source = income_sources["一、营业收入"]
            forecast.cell(row_number, 2, source[3]); _map_direct(mappings, forecast.title, forecast.cell(row_number, 2), source)
        elif label == "净利润" and income_sources.get("四、净利润"):
            source = income_sources["四、净利润"]
            forecast.cell(row_number, 2, source[3]); _map_direct(mappings, forecast.title, forecast.cell(row_number, 2), source)
        else:
            _input_fill(forecast.cell(row_number, 2)); _map_pending(mappings, forecast.title, forecast.cell(row_number, 2), "历史科目未识别或需按项目口径调整")
        for column in range(3, 8):
            _input_fill(forecast.cell(row_number, column))
            _map_pending(mappings, forecast.title, forecast.cell(row_number, column), "预测假设/经营计划")
    _style(forecast)

    for sheet_name, title, rows in (
        ("营运资本预测", "营运资本预测", ("应收账款", "存货", "预付款项", "应付账款", "合同负债", "营运资本变动")),
        ("折旧摊销及资本开支", "折旧、摊销及资本开支", ("折旧", "摊销", "资本性支出", "固定资产原值", "无形资产原值")),
    ):
        sheet = workbook.create_sheet(sheet_name)
        sheet.append([title])
        sheet.append(["项目", "预测第1年", "预测第2年", "预测第3年", "预测第4年", "预测第5年", "假设/来源"])
        for row_number, label in enumerate(rows, 3):
            sheet.cell(row_number, 1, label)
            for column in range(2, 7):
                _input_fill(sheet.cell(row_number, column)); _map_pending(mappings, sheet.title, sheet.cell(row_number, column), "预测假设或管理层预算")
        _style(sheet)

    fcf = workbook.create_sheet("企业自由现金流")
    fcf.append(["企业自由现金流测算"])
    fcf.append(["项目", "预测第1年", "预测第2年", "预测第3年", "预测第4年", "预测第5年"])
    fcf_rows = ("息税前利润", "所得税率", "税后息税前利润", "加：折旧摊销", "减：资本性支出", "减：营运资本增加", "企业自由现金流")
    for row_number, label in enumerate(fcf_rows, 3):
        fcf.cell(row_number, 1, label)
        for column in range(2, 7):
            letter = get_column_letter(column)
            if label == "税后息税前利润":
                fcf.cell(row_number, column, f'={letter}3*(1-{letter}4)')
            elif label == "企业自由现金流":
                fcf.cell(row_number, column, f'={letter}5+{letter}6-{letter}7-{letter}8')
            else:
                _input_fill(fcf.cell(row_number, column)); _map_pending(mappings, fcf.title, fcf.cell(row_number, column), "收益法预测输入")
        if label in {"税后息税前利润", "企业自由现金流"}:
            mappings.append([fcf.title, f"B{row_number}:F{row_number}", "", "", "", "", "", "公式计算", "由同列预测输入计算", "确定性公式"])
    _style(fcf)

    wacc = workbook.create_sheet("折现率WACC")
    wacc.append(["折现率（WACC）测算"])
    wacc.append(["参数", "取值", "来源/说明"])
    for row_number, label in enumerate(("无风险利率", "市场风险溢价", "贝塔系数", "特有风险调整", "税后债务成本", "权益资本成本", "资本结构（权益）", "资本结构（债务）", "WACC", "永续增长率"), 3):
        wacc.cell(row_number, 1, label)
        if label == "权益资本成本":
            wacc.cell(row_number, 2, "=B3+B4*B5+B6")
        elif label == "WACC":
            wacc.cell(row_number, 2, "=B9*B8+B10*B7")
        else:
            _input_fill(wacc.cell(row_number, 2)); _map_pending(mappings, wacc.title, wacc.cell(row_number, 2), "估值参数/市场数据")
        if label in {"权益资本成本", "WACC"}:
            mappings.append([wacc.title, f"B{row_number}", "", "", "", "", "", "公式计算", "CAPM/WACC 公式", "确定性公式"])
    _style(wacc)

    for sheet_name, title, rows in (
        ("非经营性资产负债", "溢余/非经营性资产负债", ("溢余货币资金", "非经营性资产", "非经营性负债")),
        ("付息负债明细表", "付息负债明细表", ("短期借款", "一年内到期的非流动负债", "长期借款", "租赁负债", "其他付息负债")),
    ):
        sheet = workbook.create_sheet(sheet_name)
        sheet.append([title])
        sheet.append(["项目", "账面余额", "评估调整", "评估值/扣减额", "来源/说明"])
        for row_number, label in enumerate(rows, 3):
            sheet.cell(row_number, 1, label)
            source = balance_sources.get(label)
            if source:
                sheet.cell(row_number, 2, source[3]); _map_direct(mappings, sheet.title, sheet.cell(row_number, 2), source)
            else:
                _input_fill(sheet.cell(row_number, 2)); _map_pending(mappings, sheet.title, sheet.cell(row_number, 2), "审计材料未识别或须进行经营性判断")
            _input_fill(sheet.cell(row_number, 3)); _input_fill(sheet.cell(row_number, 4))
            _map_pending(mappings, sheet.title, sheet.cell(row_number, 3), "经营性判断/评估调整")
            _map_pending(mappings, sheet.title, sheet.cell(row_number, 4), "评估师结论")
        _style(sheet)

    result = workbook.create_sheet("收益法评估结果汇总")
    result.append(["收益法评估结果汇总"])
    result.append(["项目", "金额", "计算或来源"])
    result_rows = (
        ("预测期自由现金流现值", "=SUM('企业自由现金流'!B9:F9)", "企业自由现金流表"),
        ("终值", "", "需以永续增长率或退出倍数测算"),
        ("终值现值", "", "需以折现期和WACC测算"),
        ("企业价值", "=SUM(B3:B5)", "预测现金流现值+终值现值"),
        ("加：溢余/非经营性资产", "", "非经营性资产负债表"),
        ("减：付息负债", "", "付息负债明细表"),
        ("股东全部权益价值", "=B6+B7-B8", "企业价值桥接"),
    )
    for row_number, (label, formula, note) in enumerate(result_rows, 3):
        result.cell(row_number, 1, label)
        if formula:
            result.cell(row_number, 2, formula)
            mappings.append([result.title, f"B{row_number}", "", "", "", "", "", "公式计算", formula, "确定性公式"])
        else:
            _input_fill(result.cell(row_number, 2)); _map_pending(mappings, result.title, result.cell(row_number, 2), note)
        result.cell(row_number, 3, note)
    _style(result)

    for sheet_name, title, headers in (
        ("可比公司", "市场法可比公司", ("可比公司", "证券代码", "主营业务", "PE", "PB", "EV/EBITDA", "调整说明")),
        ("市场法评估结果汇总", "市场法评估结果汇总", ("指标", "取值", "来源/说明")),
    ):
        sheet = workbook.create_sheet(sheet_name)
        sheet.append([title]); sheet.append(list(headers))
        for column in range(1, len(headers) + 1):
            _input_fill(sheet.cell(3, column)); _map_pending(mappings, sheet.title, sheet.cell(3, column), "市场数据/可比公司筛选/评估师判断")
        _style(sheet)


def build_role_workbooks(
    *,
    output_dir: Path,
    document_name: str,
    pages: list[dict[str, Any]],
) -> dict[str, Path]:
    """Create the two role workbooks and an auditable mapping graph for one report."""
    output_dir.mkdir(parents=True, exist_ok=True)
    tables = tables_from_pages(document_name, pages)
    result: dict[str, Path] = {}
    graph_lines = ["flowchart LR", f'  source["{document_name}"] --> ocr["PDF/OCR 表格"]']
    for role in (ROLE_REPORTING, ROLE_INCOME):
        role_tables = [table for table in tables if role in table.roles]
        workbook = Workbook()
        workbook.remove(workbook.active)
        _add_metadata(workbook, document_name, role, role_tables)
        used = set(workbook.sheetnames)
        mappings: list[list[Any]] = []
        for table in role_tables:
            _write_table(workbook, table, used, mappings)
            graph_lines.append(f'  ocr --> {role}_{table.page_number}["{table.category} / 第{table.page_number}页"]')
            graph_lines.append(f'  {role}_{table.page_number} --> {role}["{ROLE_LABELS[role]}"]')
        if role == ROLE_REPORTING:
            _add_asset_workpapers(workbook, role_tables, mappings)
        else:
            # The income model needs the audited balance sheet as a bridge for
            # net debt and non-operating assets, so deliberately pass both
            # statement types rather than treating the income workbook as a
            # profit-table-only export.
            _add_income_workpapers(
                workbook,
                [table for table in tables if table.category in {"资产负债表", "利润表", "现金流量表"}],
                mappings,
            )
        if not role_tables:
            sheet = workbook.create_sheet("材料缺口")
            sheet.append(["原因", "处理规则"])
            sheet.append(["本报告没有直接证明该角色所需的财务表", "保留完整工作簿结构；黄色蓝字单元格必须由评估底稿或人工资料补齐。"])
            _style(sheet)
        _add_accounting_checks(workbook, mappings)
        _add_mapping_sheet(workbook, mappings)
        path = output_dir / f"{ROLE_LABELS[role]}.xlsx"
        workbook.save(path)
        result[role] = path
    (output_dir / "映射关系图.mmd").write_text("\n".join(dict.fromkeys(graph_lines)) + "\n", encoding="utf-8")
    (output_dir / "规则摘要.json").write_text(
        json.dumps(
            {
                "source": document_name,
                "tables": [
                    {
                        "page": table.page_number,
                        "table_id": table.table_id,
                        "category": table.category,
                        "roles": list(table.roles),
                        "evidence_level": table.confidence,
                    }
                    for table in tables
                ],
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    return result
